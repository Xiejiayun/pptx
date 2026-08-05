import { createHash } from 'node:crypto';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Readable } from 'node:stream';
import { fileURLToPath } from 'node:url';
import { describe, expect, it, vi } from 'vitest';
import {
  ChartModel,
  chartWorkbookMatches,
  degrees,
  ImageModel,
  inches,
  inspectRasterImage,
  MediaModel,
  OUTPUT_TYPES,
  PRESET_SHAPE_TYPES,
  PPTX_VERSION,
  PptxDocument,
  SCHEME_COLORS,
  ShapeModel,
  TableModel,
  TEXT_ALIGNMENTS,
  TEXT_VERTICAL_ALIGNMENTS,
  ValidationError,
  type AddTableCellInput,
  type AddShapeOptions,
  type CustomGeometry,
  type OutputType,
  type PresetShapeType,
  type ShapeArrows,
  type ShapeFill,
  type ShapeLine,
  type ShapeShadow,
  type SlideMasterObject,
} from '@pptx/sdk';
import { importPptxGenJS } from './index.js';

interface BorderProps {
  readonly type?: 'none' | 'dash' | 'solid';
  readonly color?: string;
  readonly pt?: number;
}

interface PptxGenJSTableCell {
  readonly text?: string | readonly PptxGenJSTableCell[];
  readonly options?: Record<string, unknown>;
}

interface PptxGenJSChartData {
  readonly name?: string;
  readonly labels?: readonly string[] | readonly (readonly string[])[];
  readonly values?: readonly number[];
  readonly sizes?: readonly number[];
}

interface PptxGenJSSlide {
  background?: {
    readonly alpha?: number;
    readonly color?: string;
    readonly data?: string;
    readonly fill?: string;
    readonly path?: string;
    readonly src?: string;
    readonly transparency?: number;
    readonly type?: 'none' | 'solid';
  };
  color?: string | undefined;
  hidden: unknown;
  newAutoPagedSlides: PptxGenJSSlide[];
  slideNumber?: PptxGenJSSlideNumberProps;
  addChart(
    type: string | readonly {
      readonly type: string;
      readonly data: readonly PptxGenJSChartData[];
      readonly options: Record<string, unknown>;
    }[],
    data: readonly PptxGenJSChartData[] | Record<string, unknown>,
    options?: Record<string, unknown> | PptxGenJSPublicChartOptions,
  ): PptxGenJSSlide;
  addImage(options: Record<string, unknown>): void;
  addMedia(options: PptxGenJSMediaOptions): void;
  addNotes(notes: string): PptxGenJSSlide;
  addShape(type: string, options?: PptxGenJSShapeOptions): void;
  addText(
    text: string | readonly { readonly text: string; readonly options?: Record<string, unknown> }[],
    options: Record<string, unknown>,
  ): void;
  addTable(
    rows: readonly (readonly (string | PptxGenJSTableCell)[])[],
    options: Record<string, unknown>,
  ): PptxGenJSSlide;
}

interface PptxGenJSSlideNumberProps {
  readonly x?: number;
  readonly y?: number;
  readonly w?: number;
  readonly h?: number;
  readonly align?: 'left' | 'center' | 'right' | 'justify';
  readonly valign?: 'top' | 'middle' | 'bottom';
  readonly margin?: number | readonly [number, number, number, number];
  readonly fontFace?: string;
  readonly fontSize?: number;
  readonly lang?: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: string;
  readonly transparency?: number;
}

type PptxGenJSCoord = number | string;

type PptxGenJSCustomPoint =
  | { readonly x: PptxGenJSCoord; readonly y: PptxGenJSCoord; readonly moveTo?: boolean }
  | {
      readonly x: PptxGenJSCoord;
      readonly y: PptxGenJSCoord;
      readonly curve: {
        readonly type: 'arc';
        readonly hR: PptxGenJSCoord;
        readonly wR: PptxGenJSCoord;
        readonly stAng: number;
        readonly swAng: number;
      };
    }
  | {
      readonly x: PptxGenJSCoord;
      readonly y: PptxGenJSCoord;
      readonly curve: {
        readonly type: 'cubic';
        readonly x1: PptxGenJSCoord;
        readonly y1: PptxGenJSCoord;
        readonly x2: PptxGenJSCoord;
        readonly y2: PptxGenJSCoord;
      };
    }
  | {
      readonly x: PptxGenJSCoord;
      readonly y: PptxGenJSCoord;
      readonly curve: {
        readonly type: 'quadratic';
        readonly x1: PptxGenJSCoord;
        readonly y1: PptxGenJSCoord;
      };
    }
  | { readonly close: true };

interface PptxGenJSShapeOptions extends Record<string, unknown> {
  readonly points?: readonly PptxGenJSCustomPoint[];
}

type PptxGenJSPlaceholderType = 'title' | 'body' | 'pic' | 'chart' | 'tbl' | 'media';

interface PptxGenJSPlaceholderOptions extends Record<string, unknown> {
  readonly name: string;
  readonly type: PptxGenJSPlaceholderType;
  readonly x?: PptxGenJSCoord;
  readonly y?: PptxGenJSCoord;
  readonly w?: PptxGenJSCoord;
  readonly h?: PptxGenJSCoord;
}

interface PptxGenJSMasterChart {
  readonly type: string | readonly {
    readonly type: string;
    readonly data: readonly PptxGenJSChartData[];
    readonly options: Record<string, unknown>;
  }[];
  readonly data: readonly PptxGenJSChartData[] | Record<string, unknown>;
  readonly opts: Record<string, unknown>;
}

type PptxGenJSMasterObject =
  | { readonly chart: PptxGenJSMasterChart }
  | { readonly image: Readonly<Record<string, unknown>> }
  | { readonly line: Readonly<PptxGenJSShapeOptions> }
  | { readonly rect: Readonly<PptxGenJSShapeOptions> }
  | {
      readonly text: {
        readonly text: string;
        readonly options: Readonly<Record<string, unknown>>;
      };
    }
  | {
      readonly placeholder: {
        readonly text?: string;
        readonly options: Readonly<PptxGenJSPlaceholderOptions>;
      };
    };

type PptxGenJSPublicInstance = InstanceType<typeof import('pptxgenjs').default>;
type PptxGenJSPublicChartOptions = NonNullable<
  Parameters<ReturnType<PptxGenJSPublicInstance['addSlide']>['addChart']>[2]
>;
type PptxGenJSPublicShapeOptions = NonNullable<
  Parameters<ReturnType<PptxGenJSPublicInstance['addSlide']>['addShape']>[1]
>;
type PptxGenJSPublicSlideNumberOptions =
  ReturnType<PptxGenJSPublicInstance['addSlide']>['slideNumber'];
type PptxGenJSPublicTextOptions = Parameters<
  ReturnType<PptxGenJSPublicInstance['addSlide']>['addText']
>[1];
type PptxGenJSPublicText = Parameters<
  ReturnType<PptxGenJSPublicInstance['addSlide']>['addText']
>[0];
type PptxGenJSPublicTableCell = Parameters<
  ReturnType<PptxGenJSPublicInstance['addSlide']>['addTable']
>[0][number][number];
type PptxGenJSPublicTableOptions = Parameters<
  ReturnType<PptxGenJSPublicInstance['addSlide']>['addTable']
>[1];
type PptxGenJSMediaOptions = Parameters<
  ReturnType<PptxGenJSPublicInstance['addSlide']>['addMedia']
>[0];

const publicCustomShapeOptions: PptxGenJSPublicShapeOptions = {
  x: 1,
  y: 1,
  w: 4,
  h: 3,
  points: [{ x: 0, y: 0 }],
};
const publicSlideNumberOptions: PptxGenJSPublicSlideNumberOptions = {
  x: 0,
  y: 0,
  w: 1,
  h: 0.3,
  align: 'center',
  valign: 'middle',
  fontFace: 'Aptos',
  fontSize: 18,
  lang: 'zh-CN',
  bold: true,
  italic: true,
  color: 'FF3399',
  transparency: 25,
  margin: [1, 2, 3, 4],
};
const publicTextBoxOptions: PptxGenJSPublicTextOptions = { isTextBox: true };
const publicBreakLineText: PptxGenJSPublicText = [
  { text: 'First', options: { breakLine: true } },
  { text: 'Second', options: {} },
];
const publicTableSpanCell: PptxGenJSPublicTableCell = {
  text: 'Merged',
  options: { colspan: 2, rowspan: 2 },
};
const publicMeasuredTableCell: PptxGenJSPublicTableCell = {
  text: 'Weighted',
  options: { autoPageCharWeight: -1, autoPageLineWeight: 1 },
};
const publicMeasuredTableOptions: PptxGenJSPublicTableOptions = {
  autoPage: true,
  autoPageCharWeight: 0,
  autoPageLineWeight: 0,
};
const unsupportedPublicTableSpanCell: PptxGenJSPublicTableCell = {
  text: 'Invalid',
  options: {
    // @ts-expect-error PptxGenJS 4.0.1 declares numeric span fields.
    colspan: '2',
  },
};
const unsupportedPublicBreakLineText: PptxGenJSPublicText = [{
  text: 'Invalid',
  options: {
    // @ts-expect-error PptxGenJS 4.0.1 types require a boolean line-break flag.
    breakLine: 'true',
  },
}];
const unsupportedPublicTextBoxOptions: PptxGenJSPublicTextOptions = {
  // @ts-expect-error PptxGenJS 4.0.1 types require a boolean text-box flag.
  isTextBox: 'true',
};
const unsupportedPublicHandleOptions: PptxGenJSPublicShapeOptions = {
  x: 1,
  y: 1,
  w: 4,
  h: 3,
  // @ts-expect-error PptxGenJS 4.0.1 exposes no arbitrary adjustment-handle input.
  handles: [],
};
const unsupportedPublicConnectionSiteOptions: PptxGenJSPublicShapeOptions = {
  x: 1,
  y: 1,
  w: 4,
  h: 3,
  // @ts-expect-error PptxGenJS 4.0.1 exposes no arbitrary connection-site input.
  connectionSites: [],
};
const unsupportedPublicGuideOptions: PptxGenJSPublicShapeOptions = {
  x: 1,
  y: 1,
  w: 4,
  h: 3,
  // @ts-expect-error PptxGenJS 4.0.1 exposes no arbitrary guide-formula input.
  guides: [],
};
const unsupportedPublicEvaluatorOptions: PptxGenJSPublicShapeOptions = {
  x: 1,
  y: 1,
  w: 4,
  h: 3,
  // @ts-expect-error PptxGenJS 4.0.1 exposes no geometry-evaluator input.
  evaluateCustomGeometry: true,
};
void [
  publicCustomShapeOptions,
  publicSlideNumberOptions,
  publicTextBoxOptions,
  publicBreakLineText,
  publicTableSpanCell,
  publicMeasuredTableCell,
  publicMeasuredTableOptions,
  unsupportedPublicTableSpanCell,
  unsupportedPublicBreakLineText,
  unsupportedPublicTextBoxOptions,
  unsupportedPublicHandleOptions,
  unsupportedPublicConnectionSiteOptions,
  unsupportedPublicGuideOptions,
  unsupportedPublicEvaluatorOptions,
];

interface PptxGenJSInstance {
  readonly version: string;
  readonly AlignH: Readonly<Record<string, string>>;
  readonly AlignV: Readonly<Record<string, string>>;
  readonly OutputType: Readonly<Record<string, string>>;
  readonly presLayout: {
    readonly name: string;
    readonly width: number;
    readonly height: number;
  };
  readonly ChartType: Readonly<Record<string, string>>;
  readonly ShapeType: Readonly<Record<string, string>>;
  readonly SchemeColor: {
    readonly text1: 'tx1';
    readonly text2: 'tx2';
    readonly background1: 'bg1';
    readonly background2: 'bg2';
    readonly accent1: 'accent1';
    readonly accent2: 'accent2';
    readonly accent3: 'accent3';
    readonly accent4: 'accent4';
    readonly accent5: 'accent5';
    readonly accent6: 'accent6';
  };
  author: string;
  company: string;
  layout: string;
  revision: string;
  rtlMode: unknown;
  subject: string;
  theme: {
    readonly headFontFace?: string;
    readonly bodyFontFace?: string;
  } | undefined;
  title: string;
  addSection(options: { readonly title: string; readonly order?: number }): void;
  addSlide(masterName?: string): PptxGenJSSlide;
  addSlide(options?: { readonly masterName?: string; readonly sectionTitle?: string }): PptxGenJSSlide;
  defineLayout(layout: {
    readonly name: string;
    readonly width: number;
    readonly height: number;
  }): void;
  defineSlideMaster(options: {
    readonly title: string;
    readonly background?: {
      readonly alpha?: number;
      readonly color?: string;
      readonly data?: string;
      readonly fill?: string;
      readonly path?: string;
      readonly src?: string;
      readonly transparency?: number;
      readonly type?: 'none' | 'solid';
    };
    readonly bkgd?: string | {
      readonly alpha?: number;
      readonly color?: string;
      readonly data?: string;
      readonly fill?: string;
      readonly path?: string;
      readonly src?: string;
      readonly transparency?: number;
      readonly type?: 'none' | 'solid';
    };
    readonly margin?: number | readonly [number, number, number, number];
    readonly slideNumber?: PptxGenJSSlideNumberProps;
    readonly objects?: readonly PptxGenJSMasterObject[];
  }): void;
  tableToSlides(elementId: string, options?: Record<string, unknown>): void;
  stream(options?: { readonly compression?: boolean }): Promise<unknown>;
  write(options: { outputType: 'nodebuffer'; compression: boolean }): Promise<Uint8Array>;
  write(options: { outputType: 'uint8array'; compression: boolean }): Promise<Uint8Array>;
  write(options: { outputType: 'arraybuffer'; compression: boolean }): Promise<ArrayBuffer>;
  write(options: { outputType: 'base64'; compression: boolean }): Promise<string>;
  write(options: { outputType: 'binarystring'; compression: boolean }): Promise<string>;
  write(options: { outputType: 'blob'; compression: boolean }): Promise<Blob>;
  write(options: { outputType: 'STREAM'; compression: boolean }): Promise<Uint8Array>;
  write(options: { outputType: OutputType; compression: boolean }): Promise<unknown>;
  writeFile(options?: {
    readonly compression?: boolean;
    readonly fileName?: string;
  }): Promise<string>;
}

const require = createRequire(import.meta.url);
const PptxGenJS = require('pptxgenjs') as new () => PptxGenJSInstance;
const PNG_DATA_URI =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAACXBIWXMAAAABAAAAAQBPJcTWAAAAEElEQVR4nGP8ywACLGCSAQANEQED1LYyQAAAAABJRU5ErkJggg==';
const JPEG_DATA_URI =
  'data:image/jpeg;base64,/9j/4AAQSkZJRgABAgAAAQABAAD//gAQTGF2YzYyLjI4LjEwMAD/2wBDAAgEBAQEBAUFBQUFBQYGBgYGBgYGBgYGBgYHBwcICAgHBwcGBgcHCAgICAkJCQgICAgJCQoKCgwMCwsODg4RERT/xABMAAEBAAAAAAAAAAAAAAAAAAAABgEBAQAAAAAAAAAAAAAAAAAABgcQAQAAAAAAAAAAAAAAAAAAAAARAQAAAAAAAAAAAAAAAAAAAAD/wAARCAACAAIDASIAAhEAAxEA/9oADAMBAAIRAxEAPwCLAE1/f//Z';
const GIF_DATA_URI =
  'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==';
const SVG_BYTES = new TextEncoder().encode(
  '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1600 900">'
  + '<rect width="1600" height="900" fill="#4472C4"/></svg>',
);
const SVG_DATA_URI = `data:image/svg+xml;base64,${Buffer.from(SVG_BYTES).toString('base64')}`;
const SVG_EXTENSION_URI = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';
const SVG_NAMESPACE = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';
const IMAGE_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const OFFICE_MEDIA_RELATIONSHIP =
  'http://schemas.microsoft.com/office/2007/relationships/media';
const OFFICE_MEDIA_EXTENSION_URI = '{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}';
const sectionState = (document: PptxDocument) =>
  document.sections?.map(({ title, slideIds }) => ({ title, slideIds }));

interface TableToSlidesFixtureCell {
  readonly localName: 'td' | 'th';
  readonly innerText: string;
  readonly offsetWidth: number;
  readonly colSpan: number;
  readonly rowSpan: number;
  readonly style: Readonly<Record<string, string>>;
  getAttribute(name: string): string | null;
}

function tableToSlidesFixture(options: {
  readonly bodyRows?: number;
  readonly fixedFirstWidth?: string;
  readonly headerRows?: number;
} = {}) {
  const style = Object.freeze({
    color: 'rgb(1, 2, 3)',
    'background-color': 'rgb(240, 241, 242)',
    'font-family': 'Arial, sans-serif',
    'font-size': '12px',
    'font-weight': '700',
    'text-align': 'left',
    'vertical-align': 'top',
    direction: 'ltr',
    'padding-top': '1px',
    'padding-right': '1px',
    'padding-bottom': '1px',
    'padding-left': '1px',
    'border-top-style': 'solid',
    'border-right-style': 'solid',
    'border-bottom-style': 'solid',
    'border-left-style': 'solid',
    'border-top-width': '1px',
    'border-right-width': '1px',
    'border-bottom-width': '1px',
    'border-left-width': '1px',
    'border-top-color': 'rgb(4, 5, 6)',
    'border-right-color': 'rgb(4, 5, 6)',
    'border-bottom-color': 'rgb(4, 5, 6)',
    'border-left-color': 'rgb(4, 5, 6)',
  });
  const makeCell = (
    text: string,
    localName: 'td' | 'th',
    width: number,
    attributes: Readonly<Record<string, string>> = {},
  ): TableToSlidesFixtureCell => ({
    localName,
    innerText: text,
    offsetWidth: width,
    colSpan: 1,
    rowSpan: 1,
    style,
    getAttribute(name: string) {
      return Object.hasOwn(attributes, name) ? attributes[name]! : null;
    },
  });
  const headRows = Array.from({ length: options.headerRows ?? 1 }, (_, index) => ({
    cells: [
      makeCell(
        index === 0 ? 'Header A' : `Header ${index + 1} A`,
        'th',
        100,
        index === 0 && options.fixedFirstWidth !== undefined
          ? { 'data-pptx-width': options.fixedFirstWidth }
          : {},
      ),
      makeCell(index === 0 ? 'Header B' : `Header ${index + 1} B`, 'th', 100),
    ],
  }));
  const headCells = headRows[0]!.cells;
  const bodyRows = Array.from({ length: options.bodyRows ?? 2 }, (_, index) => ({
    cells: [
      makeCell(`A${index}`, 'td', 100),
      makeCell(
        `B${index} ${Array.from({
          length: options.bodyRows && options.bodyRows > 10 ? 8 : 1,
        }, () => 'content').join(' ')}`,
        'td',
        100,
      ),
    ],
  }));
  const defaultView = {
    getComputedStyle(element: unknown) {
      const values = (element as TableToSlidesFixtureCell).style;
      return {
        getPropertyValue(name: string): string {
          return values[name] ?? '';
        },
      };
    },
  };
  const table = {
    localName: 'table',
    tHead: { rows: headRows },
    tBodies: [{ rows: bodyRows }],
    tFoot: null,
    ownerDocument: undefined as unknown,
  };
  const document = {
    defaultView,
    getElementById(id: string): unknown {
      return id === 'report' ? table : null;
    },
    querySelectorAll(selector: string): readonly unknown[] {
      if (selector === '#report tr:first-child th') return headCells;
      if (selector === '#report tr:first-child td') return [];
      if (selector === '#report thead tr') return headRows;
      if (selector === '#report tbody tr') return bodyRows;
      if (selector === '#report tfoot tr') return [];
      return [];
    },
    querySelector(selector: string): unknown {
      const match = selector.match(/^#report thead tr:first-child th:nth-child\((\d+)\)$/);
      return match ? headCells[Number(match[1]) - 1] ?? null : null;
    },
  };
  table.ownerDocument = document;
  return { document, window: defaultView };
}

async function withTableToSlidesGlobals<T>(
  fixture: ReturnType<typeof tableToSlidesFixture>,
  operation: () => T | Promise<T>,
): Promise<T> {
  const previousDocument = Object.getOwnPropertyDescriptor(globalThis, 'document');
  const previousWindow = Object.getOwnPropertyDescriptor(globalThis, 'window');
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    writable: true,
    value: fixture.document,
  });
  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    writable: true,
    value: fixture.window,
  });
  try {
    return await operation();
  } finally {
    if (previousDocument) Object.defineProperty(globalThis, 'document', previousDocument);
    else Reflect.deleteProperty(globalThis, 'document');
    if (previousWindow) Object.defineProperty(globalThis, 'window', previousWindow);
    else Reflect.deleteProperty(globalThis, 'window');
  }
}

function zipCompressionMethods(bytes: Uint8Array): number[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = bytes.byteLength - 22;
  while (eocd >= 0 && view.getUint32(eocd, true) !== 0x0605_4b50) eocd -= 1;
  if (eocd < 0) throw new Error('ZIP EOCD not found');
  const entries = view.getUint16(eocd + 10, true);
  let offset = view.getUint32(eocd + 16, true);
  const methods: number[] = [];
  for (let index = 0; index < entries; index += 1) {
    if (view.getUint32(offset, true) !== 0x0201_4b50) {
      throw new Error('ZIP central directory entry not found');
    }
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const name = new TextDecoder().decode(
      bytes.subarray(offset + 46, offset + 46 + nameLength),
    );
    if (!name.endsWith('/')) methods.push(view.getUint16(offset + 10, true));
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return methods;
}

function pngHeader(width: number, height: number): Uint8Array<ArrayBuffer> {
  const bytes = new Uint8Array(24);
  bytes.set([137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82]);
  bytes[16] = width >>> 24;
  bytes[17] = width >>> 16;
  bytes[18] = width >>> 8;
  bytes[19] = width;
  bytes[20] = height >>> 24;
  bytes[21] = height >>> 16;
  bytes[22] = height >>> 8;
  bytes[23] = height;
  return bytes;
}

function pngDataUri(bytes: Uint8Array): string {
  return `data:image/png;base64,${Buffer.from(bytes).toString('base64')}`;
}

async function openPptxGenJSPublicOutput(
  presentation: PptxGenJSInstance,
): Promise<PptxDocument> {
  return PptxDocument.open(await presentation.write({
    outputType: 'nodebuffer',
    compression: true,
  }));
}

function shapeXml(document: PptxDocument, slideIndex: number, id: number): string {
  const slide = document.slides[slideIndex];
  if (!slide) throw new Error(`Slide ${slideIndex} was not found`);
  const xml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
  const shapes = [...xml.matchAll(/<p:sp(?:\s[^>]*)?>[\s\S]*?<\/p:sp>/g)]
    .map((match) => match[0]);
  const shape = shapes.find((candidate) => new RegExp(
    `<p:cNvPr\\b[^>]*\\bid="${id}"(?:\\s|/|>)`,
  ).test(candidate));
  if (!shape) throw new Error(`Shape ${id} was not found on slide ${slideIndex}`);
  return shape;
}

function pictureXml(document: PptxDocument, slideIndex: number, id: number): string {
  const slide = document.slides[slideIndex];
  if (!slide) throw new Error(`Slide ${slideIndex} was not found`);
  const xml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
  const pictures = [...xml.matchAll(/<p:pic(?:\s[^>]*)?>[\s\S]*?<\/p:pic>/g)]
    .map((match) => match[0]);
  const picture = pictures.find((candidate) => new RegExp(
    `<p:cNvPr\\b[^>]*\\bid="${id}"(?:\\s|/|>)`,
  ).test(candidate));
  if (!picture) throw new Error(`Image ${id} was not found on slide ${slideIndex}`);
  return picture;
}

function slideXml(document: PptxDocument, slideIndex: number): string {
  const slide = document.slides[slideIndex];
  if (!slide) throw new Error(`Slide ${slideIndex} was not found`);
  return new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
}

function slideNumberOwnerState(document: PptxDocument, partUri: string) {
  const xml = new TextDecoder().decode(document.opcPackage.requirePart(partUri).bytes);
  const shapes = [...xml.matchAll(/<p:sp(?:\s[^>]*)?>[\s\S]*?<\/p:sp>/g)]
    .map((match) => match[0])
    .filter((shape) => /<p:ph\b[^>]*\btype="sldNum"/.test(shape));
  const shape = shapes[0];
  const shapeId = shape?.match(/<p:cNvPr\b[^>]*\bid="([0-9]+)"/)?.[1];
  const field = shape?.match(/<a:fld\b[^>]*\btype="slidenum"[^>]*>[\s\S]*?<\/a:fld>/)?.[0];
  const cache = field?.match(/<a:t>([\s\S]*?)<\/a:t>/)?.[1];
  const ids = [...xml.matchAll(/<p:cNvPr\b[^>]*\bid="([0-9]+)"/g)]
    .map((match) => match[1]);
  const masterFlag = xml.match(/<p:hf\b[^>]*\bsldNum="([^"]+)"/)?.[1];
  return {
    ownerCount: shapes.length,
    shapeId: shapeId === undefined ? undefined : Number(shapeId),
    shapeIdOccurrences: shapeId === undefined
      ? 0
      : ids.filter((id) => id === shapeId).length,
    fieldType: field?.match(/\btype="([^"]+)"/)?.[1],
    cache,
    masterFlag,
    xml,
  };
}

function slideBackgroundStructuralState(document: PptxDocument, slideIndex: number) {
  const slide = document.slides[slideIndex];
  if (!slide) throw new Error(`Slide ${slideIndex} was not found`);
  const xml = slideXml(document, slideIndex);
  const directBackground = xml.match(/<p:bg(?:\s[^>]*)?>[\s\S]*?<\/p:bg>/)?.[0];
  const background = slide.background;
  const relationship = background?.kind === 'image'
    ? slide.relationships.find(({ type }) => type === IMAGE_RELATIONSHIP)
    : undefined;
  const part = relationship?.resolvedTarget
    ? document.opcPackage.requirePart(relationship.resolvedTarget)
    : undefined;
  return {
    kind: background?.kind,
    color: background?.kind === 'solid' ? background.color : undefined,
    transparency: background?.kind === 'solid' ? background.transparency : undefined,
    direct: {
      present: directBackground !== undefined,
      noFill: directBackground?.includes('<a:noFill/>') ?? false,
      solidFill: directBackground?.includes('<a:solidFill>') ?? false,
      blipFill: directBackground?.includes('<a:blipFill') ?? false,
      stretchFillRect: /<a:stretch>\s*<a:fillRect\s*\/>\s*<\/a:stretch>/.test(
        directBackground ?? '',
      ),
    },
    image: part && relationship
      ? {
          relationshipType: relationship.type,
          targetMode: relationship.targetMode,
          contentType: part.contentType,
          extension: relationship.resolvedTarget?.match(/\.[^/.]+$/)?.[0],
          sha256: createHash('sha256').update(part.bytes).digest('hex'),
        }
      : undefined,
  };
}

function embeddedRasterState(
  document: PptxDocument,
  slideIndex: number,
  image: ImageModel,
) {
  const slide = document.slides[slideIndex]!;
  const sourcePartUri = image.sourcePartUri!;
  const part = document.opcPackage.requirePart(sourcePartUri);
  const xml = pictureXml(document, slideIndex, image.id);
  const embeddedRelationshipId = xml.match(/\br:embed="([^"]+)"/)?.[1];
  const relationship = slide.relationships.find(
    ({ id }) => id === embeddedRelationshipId,
  );
  return {
    kind: image.kind,
    name: image.name,
    altText: xml.match(/<p:cNvPr\b[^>]*\bdescr="([^"]*)"/)?.[1],
    transform: image.transform,
    contentType: part.contentType,
    bytes: [...part.bytes],
    relationshipType: relationship?.type,
    targetMode: relationship?.targetMode,
    embedded: relationship?.type.endsWith('/image') === true
      && relationship.resolvedTarget === sourcePartUri,
    sourceRectangle: image.sourceRectangle,
    sourceRectangleRaw: directSourceRectangleRaw(xml),
    rectGeometry: xml.includes('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'),
    noChangeAspect: xml.includes('<a:picLocks noChangeAspect="1"/>'),
    stretchFill: /<a:stretch>\s*<a:fillRect\/>\s*<\/a:stretch>/.test(xml),
  };
}

function embeddedRasterSizingState(
  document: PptxDocument,
  slideIndex: number,
  image: ImageModel,
) {
  const { stretchFill: _stretchSyntax, ...state } = embeddedRasterState(
    document,
    slideIndex,
    image,
  );
  return state;
}

function embeddedSvgState(
  document: PptxDocument,
  slideIndex: number,
  image: ImageModel,
) {
  const slide = document.slides[slideIndex]!;
  const xml = pictureXml(document, slideIndex, image.id);
  const fallbackRelationshipId = xml.match(
    /<a:blip\b[^>]*\br:embed="([^"]+)"/,
  )?.[1];
  const svgRelationshipId = xml.match(
    /<asvg:svgBlip\b[^>]*\br:embed="([^"]+)"/,
  )?.[1];
  const fallbackRelationship = slide.relationships.find(
    ({ id }) => id === fallbackRelationshipId,
  );
  const svgRelationship = slide.relationships.find(
    ({ id }) => id === svgRelationshipId,
  );
  const fallbackPartUri = image.fallbackPartUri!;
  const svgPartUri = image.svgPartUri!;
  const fallbackPart = document.opcPackage.requirePart(fallbackPartUri);
  const svgPart = document.opcPackage.requirePart(svgPartUri);
  return {
    kind: image.kind,
    isSvg: image.isSvg,
    name: image.name,
    altText: xml.match(/<p:cNvPr\b[^>]*\bdescr="([^"]*)"/)?.[1],
    transform: image.transform,
    sourceRectangle: image.sourceRectangle,
    sourceRectangleRaw: directSourceRectangleRaw(xml),
    fallbackContentType: fallbackPart.contentType,
    svgContentType: svgPart.contentType,
    svgBytes: [...svgPart.bytes],
    extensionUri: xml.match(
      /<a:ext\b[^>]*\buri="([^"]+)"[^>]*>[\s\S]*?<asvg:svgBlip/,
    )?.[1],
    extensionNamespace: xml.match(
      /<asvg:svgBlip\b[^>]*\bxmlns:asvg="([^"]+)"/,
    )?.[1],
    relationshipRoles: {
      fallback: {
        type: fallbackRelationship?.type,
        targetMode: fallbackRelationship?.targetMode,
        resolvesToExpectedPart: fallbackRelationship?.resolvedTarget === fallbackPartUri,
      },
      svg: {
        type: svgRelationship?.type,
        targetMode: svgRelationship?.targetMode,
        resolvesToExpectedPart: svgRelationship?.resolvedTarget === svgPartUri,
      },
    },
  };
}

function directSourceRectangleRaw(xml: string) {
  const attributes = xml.match(/<a:srcRect\b([^>]*)\/>/)?.[1];
  if (attributes === undefined) return undefined;
  const read = (name: 'l' | 't' | 'r' | 'b'): number => {
    const lexical = attributes.match(new RegExp(`(?:^|\\s)${name}="(-?[0-9]+)"`))?.[1];
    if (lexical === undefined) throw new Error(`Source rectangle ${name} was not found`);
    return Number(lexical);
  };
  return { left: read('l'), top: read('t'), right: read('r'), bottom: read('b') };
}

function packageState(document: PptxDocument) {
  return {
    parts: document.opcPackage.parts.map(({ uri, contentType, bytes, relationships }) => ({
      uri,
      contentType,
      bytes: [...bytes],
      relationships,
    })),
    graph: document.opcPackage.graph,
    mutations: [...document.opcPackage.mutations],
  };
}

interface EmbeddedMediaState {
  readonly shapeId: number;
  readonly fileElement: 'audioFile' | 'videoFile';
  readonly nameXml: string;
  readonly transform: Readonly<{
    x: number;
    y: number;
    width: number;
    height: number;
  }>;
  readonly clickAction: boolean;
  readonly aspectLocked: boolean;
  readonly officeMediaExtension: boolean;
  readonly mediaPartUri: string;
  readonly mediaContentType: string;
  readonly mediaExtension: string;
  readonly mediaBytes: readonly number[];
  readonly posterPartUri: string;
  readonly posterContentType: string;
  readonly posterExtension: string;
  readonly posterBytes: readonly number[];
  readonly relationshipRoles: Readonly<{
    kind: Readonly<{
      type: string | undefined;
      targetMode: string | undefined;
      resolvesToMedia: boolean;
    }>;
    media: Readonly<{
      type: string | undefined;
      targetMode: string | undefined;
      resolvesToMedia: boolean;
    }>;
    poster: Readonly<{
      type: string | undefined;
      targetMode: string | undefined;
      resolvesToPoster: boolean;
    }>;
  }>;
}

function embeddedMediaStates(
  document: PptxDocument,
  slideIndex: number,
): EmbeddedMediaState[] {
  const slide = document.slides[slideIndex]!;
  const xml = new TextDecoder().decode(
    document.opcPackage.requirePart(slide.partUri).bytes,
  );
  return [...xml.matchAll(/<p:pic(?:\s[^>]*)?>[\s\S]*?<\/p:pic>/g)]
    .map((match) => match[0]!)
    .filter((picture) => /<a:(?:audioFile|videoFile)\b/.test(picture))
    .map((picture) => {
      const properties = picture.match(/<p:cNvPr\b([^>]*)>/)?.[1];
      const shapeId = Number(properties?.match(/\bid="([0-9]+)"/)?.[1]);
      const nameXml = properties?.match(/\bname="([^"]*)"/)?.[1];
      const file = picture.match(/<a:(audioFile|videoFile)\b[^>]*\br:link="([^"]+)"/);
      const mediaRelationshipId = picture.match(
        /<p14:media\b[^>]*\br:embed="([^"]+)"/,
      )?.[1];
      const posterRelationshipId = picture.match(
        /<a:blip\b[^>]*\br:embed="([^"]+)"/,
      )?.[1];
      const offset = picture.match(/<a:off\b[^>]*\bx="(-?[0-9]+)"[^>]*\by="(-?[0-9]+)"/);
      const extent = picture.match(/<a:ext\b[^>]*\bcx="([0-9]+)"[^>]*\bcy="([0-9]+)"/);
      if (!properties || !Number.isSafeInteger(shapeId) || nameXml === undefined || !file || !offset || !extent) {
        throw new Error('Embedded media picture is malformed');
      }
      const kindRelationship = slide.relationships.find(({ id }) => id === file[2]);
      const mediaRelationship = slide.relationships.find(({ id }) => id === mediaRelationshipId);
      const posterRelationship = slide.relationships.find(({ id }) => id === posterRelationshipId);
      const mediaPartUri = mediaRelationship?.resolvedTarget ?? kindRelationship?.resolvedTarget;
      const posterPartUri = posterRelationship?.resolvedTarget;
      if (!mediaPartUri || !posterPartUri) throw new Error('Embedded media relationship target is missing');
      const mediaPart = document.opcPackage.requirePart(mediaPartUri);
      const posterPart = document.opcPackage.requirePart(posterPartUri);
      return {
        shapeId,
        fileElement: file[1] as 'audioFile' | 'videoFile',
        nameXml,
        transform: {
          x: Number(offset[1]),
          y: Number(offset[2]),
          width: Number(extent[1]),
          height: Number(extent[2]),
        },
        clickAction: picture.includes('action="ppaction://media"'),
        aspectLocked: picture.includes('<a:picLocks noChangeAspect="1"/>'),
        officeMediaExtension: picture.includes(
          `<p:ext uri="${OFFICE_MEDIA_EXTENSION_URI}">`,
        ),
        mediaPartUri,
        mediaContentType: mediaPart.contentType,
        mediaExtension: mediaPartUri.slice(mediaPartUri.lastIndexOf('.')),
        mediaBytes: [...mediaPart.bytes],
        posterPartUri,
        posterContentType: posterPart.contentType,
        posterExtension: posterPartUri.slice(posterPartUri.lastIndexOf('.')),
        posterBytes: [...posterPart.bytes],
        relationshipRoles: {
          kind: {
            type: kindRelationship?.type,
            targetMode: kindRelationship?.targetMode,
            resolvesToMedia: kindRelationship?.resolvedTarget === mediaPartUri,
          },
          media: {
            type: mediaRelationship?.type,
            targetMode: mediaRelationship?.targetMode,
            resolvesToMedia: mediaRelationship?.resolvedTarget === mediaPartUri,
          },
          poster: {
            type: posterRelationship?.type,
            targetMode: posterRelationship?.targetMode,
            resolvesToPoster: posterRelationship?.resolvedTarget === posterPartUri,
          },
        },
      };
    });
}

function comparableEmbeddedMediaState(state: EmbeddedMediaState) {
  return {
    nameXml: state.nameXml,
    transform: state.transform,
    clickAction: state.clickAction,
    aspectLocked: state.aspectLocked,
    officeMediaExtension: state.officeMediaExtension,
    mediaExtension: state.mediaExtension,
    mediaBytes: state.mediaBytes,
    posterContentType: state.posterContentType,
    posterExtension: state.posterExtension,
    posterBytes: state.posterBytes,
  };
}

function externalVideoState(document: PptxDocument, slideIndex: number) {
  const slide = document.slides[slideIndex]!;
  const model = slide.media.find(({ externalUrl }) => externalUrl !== undefined);
  if (!model?.externalUrl || !model.posterPartUri) {
    throw new Error('External video state was not found');
  }
  const xml = new TextDecoder().decode(
    document.opcPackage.requirePart(slide.partUri).bytes,
  );
  const picture = [...xml.matchAll(/<p:pic(?:\s[^>]*)?>[\s\S]*?<\/p:pic>/g)]
    .map((match) => match[0]!)
    .find((candidate) => candidate.includes(`<p:cNvPr id="${model.shapeId}"`));
  if (!picture) throw new Error(`External video picture ${model.shapeId} was not found`);
  const videoRelationshipId = picture.match(
    /<a:videoFile\b[^>]*\br:link="([^"]+)"/,
  )?.[1];
  const posterRelationshipId = picture.match(
    /<a:blip\b[^>]*\br:embed="([^"]+)"/,
  )?.[1];
  const videoRelationship = slide.relationships.find(({ id }) => id === videoRelationshipId);
  const posterRelationship = slide.relationships.find(({ id }) => id === posterRelationshipId);
  const poster = document.opcPackage.requirePart(model.posterPartUri);
  return {
    name: model.name,
    transform: model.transform,
    settings: model.settings,
    externalUrl: model.externalUrl,
    mediaPartUri: model.mediaPartUri,
    poster: {
      contentType: poster.contentType,
      extension: model.posterPartUri.slice(model.posterPartUri.lastIndexOf('.')),
      bytes: [...poster.bytes],
    },
    relationships: {
      video: {
        type: videoRelationship?.type,
        target: videoRelationship?.target,
        targetMode: videoRelationship?.targetMode,
      },
      poster: {
        type: posterRelationship?.type,
        targetMode: posterRelationship?.targetMode,
        resolvesToPoster: posterRelationship?.resolvedTarget === model.posterPartUri,
      },
    },
    clickAction: picture.includes('action="ppaction://media"'),
    aspectLocked: picture.includes('<a:picLocks noChangeAspect="1"/>'),
    playbackExtension: picture.includes('urn:pptx-ooxml:media'),
    timing: xml.includes('<p:timing>'),
  };
}

function comparableExternalVideoState(state: ReturnType<typeof externalVideoState>) {
  return {
    name: state.name,
    transform: state.transform,
    externalUrl: state.externalUrl,
    mediaPartUri: state.mediaPartUri,
    poster: state.poster,
    relationships: state.relationships,
  };
}

function slideRelationshipsXml(document: PptxDocument, slideIndex: number): string {
  const slidePartUri = document.slides[slideIndex]!.partUri;
  const separator = slidePartUri.lastIndexOf('/');
  const relationshipsPartUri = `${slidePartUri.slice(0, separator)}/_rels/${
    slidePartUri.slice(separator + 1)
  }.rels`;
  return new TextDecoder().decode(
    document.opcPackage.requirePart(relationshipsPartUri).bytes,
  );
}

function directShapePaintState(xml: string): { fill: 'none'; line: 'empty' } {
  const properties = xml.match(/<p:spPr(?:\s[^>]*)?>([\s\S]*?)<\/p:spPr>/)?.[1];
  if (!properties) throw new Error('Shape properties were not found');
  const afterGeometry = properties.match(/<\/a:prstGeom>([\s\S]*)$/)?.[1];
  if (afterGeometry === undefined) throw new Error('Preset geometry was not found');
  if (!/^<a:noFill\/><a:ln(?:\/>|><\/a:ln>)/.test(afterGeometry)) {
    throw new Error('Expected direct no-fill and empty line state');
  }
  return { fill: 'none', line: 'empty' };
}

function directTextPresetGeometryState(xml: string) {
  const properties = xml.match(/<p:spPr(?:\s[^>]*)?>([\s\S]*?)<\/p:spPr>/)?.[1];
  if (!properties) throw new Error('Text shape properties were not found');
  const geometries = [...properties.matchAll(
    /<a:prstGeom\b([^>]*)>([\s\S]*?)<\/a:prstGeom>/g,
  )];
  if (geometries.length !== 1) throw new Error('Expected one direct preset geometry');
  const geometryAttributes = geometries[0]?.[1];
  const geometryContents = geometries[0]?.[2];
  if (geometryAttributes === undefined || geometryContents === undefined) {
    throw new Error('Preset geometry content was not found');
  }
  const type = geometryAttributes.match(/(?:^|\s)prst="([^"]+)"/)?.[1];
  if (type === undefined) throw new Error('Preset geometry type was not found');
  const adjustmentLists = [...geometryContents.matchAll(
    /<a:avLst(?:\s[^>]*)?(?:\/>|>([\s\S]*?)<\/a:avLst>)/g,
  )];
  if (adjustmentLists.length !== 1) throw new Error('Expected one adjustment list');
  const adjustmentXml = adjustmentLists[0]![1] ?? '';
  const adjustments = [...adjustmentXml.matchAll(/<a:gd\b([^>]*)\/>/g)].map((match) => ({
    name: match[1]!.match(/(?:^|\s)name="([^"]+)"/)?.[1],
    formula: match[1]!.match(/(?:^|\s)fmla="([^"]+)"/)?.[1],
  }));
  if (adjustmentXml.replace(/<a:gd\b[^>]*\/>/g, '').trim() !== '') {
    throw new Error('Unexpected preset geometry adjustment content');
  }
  return { type, adjustments };
}

describe('importPptxGenJS', () => {
  it('reports each library runtime version through its public instance', async () => {
    const generated = new PptxGenJS();
    const native = PptxDocument.create();

    expect(generated.version).toBe('4.0.1');
    expect(native.version).toBe(PPTX_VERSION);
    expect(generated.version).not.toBe(native.version);

    await generated.write({ outputType: 'uint8array', compression: false });
    await native.write();
    expect(generated.version).toBe('4.0.1');
    expect(native.version).toBe(PPTX_VERSION);
  });

  it('matches PptxGenJS shape and text percentage coordinate output with explicit native units', async () => {
    const generated = new PptxGenJS();
    generated.defineLayout({ name: 'PERCENT_AUDIT', width: 10, height: 8 });
    generated.layout = 'PERCENT_AUDIT';
    const generatedSlide = generated.addSlide();
    generatedSlide.addShape(generated.ShapeType.rect!, {
      x: '10%',
      y: '20%',
      w: '30%',
      h: '40%',
    });
    generatedSlide.addText('Percentage text', {
      x: '12.5%',
      y: '25%',
      w: '37.5%',
      h: '50%',
    });

    const imported = await openPptxGenJSPublicOutput(generated);
    const native = PptxDocument.create({
      slideSize: { width: inches(10), height: inches(8) },
    });
    const nativeSlide = native.addSlide();
    nativeSlide.addShape('rect', {
      x: inches(1),
      y: '20%',
      width: '30%',
      height: '40%',
    });
    nativeSlide.addText('Percentage text', {
      x: '12.5%',
      y: '25%',
      width: '37.5%',
      height: '50%',
    });

    const expected = [
      {
        x: 914_400,
        y: 1_463_040,
        width: 2_743_200,
        height: 2_926_080,
        rotation: 0,
        flipHorizontal: false,
        flipVertical: false,
      },
      {
        x: 1_143_000,
        y: 1_828_800,
        width: 3_429_000,
        height: 3_657_600,
        rotation: 0,
        flipHorizontal: false,
        flipVertical: false,
      },
    ];
    expect(imported.slides[0]!.shapes.map(({ transform }) => transform)).toEqual(expected);
    expect(native.slides[0]!.shapes.map(({ transform }) => transform)).toEqual(expected);

    await native.write({ compatibility: 'powerpoint-2010' });
    expect(native.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
    const reopened = await PptxDocument.open(await native.write());
    expect(reopened.slides[0]!.shapes.map(({ transform }) => transform)).toEqual(expected);
    expect(reopened.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('matches PptxGenJS image percentage coordinate output with explicit native units', async () => {
    const generated = new PptxGenJS();
    generated.defineLayout({ name: 'IMAGE_PERCENT_AUDIT', width: 10, height: 8 });
    generated.layout = 'IMAGE_PERCENT_AUDIT';
    generated.addSlide().addImage({
      data: PNG_DATA_URI,
      x: '10%',
      y: '20%',
      w: '30%',
      h: '40%',
    });

    const imported = await openPptxGenJSPublicOutput(generated);
    const native = PptxDocument.create({
      slideSize: { width: inches(10), height: inches(8) },
    });
    native.addSlide();
    await native.addImage(0, PNG_DATA_URI, {
      x: inches(1),
      y: '20%',
      width: '30%',
      height: '40%',
    });

    const expected = {
      x: 914_400,
      y: 1_463_040,
      width: 2_743_200,
      height: 2_926_080,
      rotation: 0,
      flipHorizontal: false,
      flipVertical: false,
    };
    expect(imported.slides[0]!.shapes[0]!.transform).toEqual(expected);
    expect(native.slides[0]!.shapes[0]!.transform).toEqual(expected);

    await native.write({ compatibility: 'powerpoint-2010' });
    expect(native.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
    const reopened = await PptxDocument.open(await native.write());
    expect(reopened.slides[0]!.shapes[0]!.transform).toEqual(expected);
    expect(reopened.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('matches legal tableToSlides rows, styles, widths, and additions', async () => {
    const generated = new PptxGenJS();
    const legacyOptions: Record<string, unknown> = {
      autoPage: true,
      addImage: {
        image: { data: PNG_DATA_URI },
        options: { x: 0.1, y: 0.1, w: 0.2, h: 0.2 },
      },
      addShape: {
        shapeName: 'rect',
        options: { x: 0.3, y: 0.3, w: 0.2, h: 0.2 },
      },
      addTable: { rows: [['K', 'V']], options: { x: 0.5, y: 0.5, w: 1 } },
      addText: { text: 'Details', options: { x: 0.5, y: 1, w: 1, h: 0.2 } },
    };
    const legacyReturn = await withTableToSlidesGlobals(tableToSlidesFixture(), () =>
      generated.tableToSlides('report', legacyOptions));
    const imported = await importPptxGenJS(generated);

    const native = PptxDocument.create();
    const nativeOptions = {
      autoPage: true,
      addImage: {
        source: PNG_DATA_URI,
        options: { x: inches(0.1), y: inches(0.1), width: inches(0.2), height: inches(0.2) },
      },
      addShape: {
        type: 'rect' as const,
        options: { x: inches(0.3), y: inches(0.3), width: inches(0.2), height: inches(0.2) },
      },
      addTable: {
        rows: [['K', 'V']],
        options: { x: inches(0.5), y: inches(0.5), width: inches(1) },
      },
      addText: {
        text: 'Details',
        options: { x: inches(0.5), y: inches(1), width: inches(1), height: inches(0.2) },
      },
    };
    const nativeBefore = structuredClone(nativeOptions);
    let returnedPromise = false;
    const nativePages = await withTableToSlidesGlobals(tableToSlidesFixture(), () => {
      const result = native.tableToSlides('report', nativeOptions);
      returnedPromise = result instanceof Promise;
      return result;
    });

    expect(legacyReturn).toBeUndefined();
    expect(returnedPromise).toBe(true);
    expect(legacyOptions).toHaveProperty('_arrObjTabHeadRows');
    expect(legacyOptions).toHaveProperty('colW');
    expect(nativeOptions).toEqual(nativeBefore);
    expect(imported.slides).toHaveLength(1);
    expect(nativePages).toHaveLength(1);
    expect(imported.slides[0]!.shapes.map(({ kind }) => kind))
      .toEqual(['table', 'image', 'shape', 'table', 'text']);
    expect(nativePages[0]!.shapes.map(({ kind }) => kind))
      .toEqual(['table', 'image', 'shape', 'table', 'text']);

    const legacyTable = imported.slides[0]!.shapes[0] as TableModel;
    const nativeTable = nativePages[0]!.shapes[0] as TableModel;
    expect(legacyTable.rows.map((row) => row.cells.map(({ text }) => text)))
      .toEqual(nativeTable.rows.map((row) => row.cells.map(({ text }) => text)));
    expect(legacyTable.columnWidths).toEqual(nativeTable.columnWidths);
    expect(legacyTable.rows[0]!.cells[0]!.fill).toEqual(nativeTable.rows[0]!.cells[0]!.fill);
    expect(legacyTable.rows[0]!.cells[0]!.margins)
      .toEqual(nativeTable.rows[0]!.cells[0]!.margins);
    expect(legacyTable.rows[0]!.cells[0]!.horizontalAlignment ?? 'left')
      .toBe(nativeTable.rows[0]!.cells[0]!.horizontalAlignment ?? 'left');
    expect(legacyTable.rows[0]!.cells[0]!.verticalAlignment ?? 'top')
      .toBe(nativeTable.rows[0]!.cells[0]!.verticalAlignment ?? 'top');
  });

  it('isolates the ignored tableToSlides fill declaration from computed CSS backgrounds', async () => {
    const generate = async (fill?: Record<string, unknown>) => {
      const generated = new PptxGenJS();
      await withTableToSlidesGlobals(tableToSlidesFixture(), () =>
        generated.tableToSlides('report', fill === undefined ? {} : { fill }));
      return openPptxGenJSPublicOutput(generated);
    };
    const omitted = await generate();
    const declared = await generate({ color: 'FF0000', transparency: 25 });
    const fills = (document: PptxDocument) => {
      const table = document.slides[0]!.shapes[0] as TableModel;
      return table.rows.map((row) => row.cells.map(({ fill }) => fill));
    };
    const expected = Array.from({ length: 3 }, () => Array(2).fill({
      kind: 'solid',
      color: { kind: 'srgb', value: 'F0F1F2' },
    }));
    expect(fills(omitted)).toEqual(expected);
    expect(fills(declared)).toEqual(expected);
    const declaredXml = new TextDecoder().decode(
      declared.opcPackage.requirePart(declared.slides[0]!.partUri).bytes,
    );
    expect(declaredXml).not.toContain('FF0000');
    expect(declaredXml).not.toContain('75000');

    const native = PptxDocument.create();
    const pages = await withTableToSlidesGlobals(tableToSlidesFixture(), () =>
      native.tableToSlides('report'));
    expect(pages).toHaveLength(1);
    expect(fills(native)).toEqual(expected);
  });

  it('locks strict tableToSlides differences from PptxGenJS 4.0.1 defects', async () => {
    const generated = new PptxGenJS();
    const legacyOptions: Record<string, unknown> = {
      autoPage: false,
      addText: { text: 'Legacy' },
    };
    await withTableToSlidesGlobals(tableToSlidesFixture({
      bodyRows: 45,
      fixedFirstWidth: '0.5',
    }), () => generated.tableToSlides('report', legacyOptions));
    const imported = await importPptxGenJS(generated);

    const native = PptxDocument.create();
    const nativePages = await withTableToSlidesGlobals(tableToSlidesFixture({
      bodyRows: 45,
      fixedFirstWidth: '0.5',
    }), () => native.tableToSlides('report', {
      autoPage: false,
      addText: { text: 'Native' },
    }));
    expect(imported.slides.length).toBeGreaterThan(1);
    expect(nativePages).toHaveLength(1);
    const legacyTable = imported.slides[0]!.shapes[0] as TableModel;
    const nativeTable = nativePages[0]!.shapes[0] as TableModel;
    expect(legacyTable.columnWidths![0]).toBe(inches(4.5));
    expect(nativeTable.columnWidths![0]).toBe(inches(0.5));
    expect(legacyOptions).not.toEqual({ autoPage: false, addText: { text: 'Legacy' } });

    const permissive = new PptxGenJS();
    await expect(withTableToSlidesGlobals(tableToSlidesFixture(), () =>
      permissive.tableToSlides('report', { autoPage: 'false' })))
      .resolves.toBeUndefined();
    const strict = PptxDocument.create();
    await expect(withTableToSlidesGlobals(tableToSlidesFixture(), () =>
      strict.tableToSlides('report', { autoPage: 'false' as never })))
      .rejects.toThrow(/autoPage.*boolean/i);
    expect(strict.slides).toEqual([]);
  });

  it('locks every declared tableToSlides option and nested addition against PptxGenJS 4.0.1', async () => {
    const propertyId = (property: string) =>
      `interface:TableToSlidesProps@property:${property}`;
    const atomGroups = {
      supported: [
        'autoPageCharWeight',
        'autoPageLineWeight',
        'autoPageRepeatHeader',
        'masterSlideName',
      ].map(propertyId),
      deliberate: [
        'class:PptxGenJS#tableToSlides',
        ...[
          'addImage',
          'addShape',
          'addTable',
          'addText',
          'autoPageSlideStartY',
          'h',
          'slideMargin',
          'verbose',
          'w',
          'x',
          'y',
        ].map(propertyId),
        'inline:interface:TableToSlidesProps@property:addImage@property:addImage.image',
        'inline:interface:TableToSlidesProps@property:addImage@property:addImage.options',
        'inline:interface:TableToSlidesProps@property:addShape@property:addShape.options',
        'inline:interface:TableToSlidesProps@property:addShape@property:addShape.shapeName',
        'inline:interface:TableToSlidesProps@property:addTable@property:addTable.options',
        'inline:interface:TableToSlidesProps@property:addTable@property:addTable.rows',
        'inline:interface:TableToSlidesProps@property:addText@property:addText.options',
        'inline:interface:TableToSlidesProps@property:addText@property:addText.text',
      ],
      deprecated: ['addHeaderToEach', 'newSlideStartY'].map(propertyId),
      defect: [
        ...[
          'align',
          'autoPage',
          'autoPageHeaderRows',
          'border',
          'colW',
          'margin',
          'objectName',
          'rowH',
          'transparency',
          'valign',
        ].map(propertyId),
        'union:interface:TableToSlidesProps@property:border#BorderProps',
        'union:interface:TableToSlidesProps@property:border#[BorderProps,BorderProps,BorderProps,BorderProps]',
        'union:interface:TableToSlidesProps@property:colW#number',
        'union:interface:TableToSlidesProps@property:colW#number[]',
        'union:interface:TableToSlidesProps@property:rowH#number',
        'union:interface:TableToSlidesProps@property:rowH#number[]',
      ],
    } as const;
    expect(Object.fromEntries(Object.entries(atomGroups).map(([status, ids]) => [
      status,
      ids.length,
    ]))).toEqual({ supported: 4, deliberate: 20, deprecated: 2, defect: 16 });
    expect(new Set(Object.values(atomGroups).flat()).size).toBe(42);

    const tableOn = (slide: PptxDocument['slides'][number]) =>
      slide.shapes.find((shape): shape is TableModel => shape instanceof TableModel)!;
    const tableState = (table: TableModel) => ({
      transform: table.transform,
      columnWidths: table.columnWidths,
      rowHeights: table.rowHeights,
      rows: table.rows.map((row) => row.cells.map((cell) => ({
        text: cell.text,
        fill: cell.fill,
        borders: cell.borders,
        margins: cell.margins,
        horizontalAlignment: cell.horizontalAlignment,
        verticalAlignment: cell.verticalAlignment,
      }))),
    });

    const active = new PptxGenJS();
    active.defineSlideMaster({ title: 'TTS-AUDIT-MASTER', objects: [] });
    const activeOptions: Record<string, unknown> = {
      masterSlideName: 'TTS-AUDIT-MASTER',
      autoPageCharWeight: -0.25,
      autoPageLineWeight: 0.25,
      autoPageRepeatHeader: true,
      autoPageHeaderRows: 1,
      autoPageSlideStartY: 1.25,
      slideMargin: [0.5, 0.5, 0.5, 0.5],
      x: 1.25,
      y: 3,
      w: 9,
      h: 6,
      addImage: {
        image: { data: PNG_DATA_URI },
        options: { x: 0.1, y: 0.1, w: 0.2, h: 0.2 },
      },
      addShape: {
        shapeName: 'rect',
        options: { x: 0.3, y: 0.3, w: 0.2, h: 0.2 },
      },
      addTable: {
        rows: [['Nested K', 'Nested V']],
        options: { x: 0.5, y: 0.5, w: 1 },
      },
      addText: {
        text: [{ text: 'Nested text', options: { bold: true } }],
        options: { x: 0.5, y: 1, w: 1, h: 0.2 },
      },
    };
    const activeReturn = await withTableToSlidesGlobals(tableToSlidesFixture({
      bodyRows: 45,
      headerRows: 2,
    }), () => active.tableToSlides('report', activeOptions));
    expect(activeReturn).toBeUndefined();
    expect(activeOptions).toHaveProperty('_arrObjTabHeadRows');
    expect(activeOptions).toHaveProperty('colW');
    const activeImported = await openPptxGenJSPublicOutput(active);
    expect(activeImported.slides.length).toBeGreaterThan(1);
    const activeLayout = activeImported.layouts.find(
      ({ name }) => name === 'TTS-AUDIT-MASTER',
    )!;
    expect(activeLayout).toBeDefined();
    for (const [index, slide] of activeImported.slides.entries()) {
      expect(slide.shapes.map(({ kind }) => kind))
        .toEqual(['table', 'image', 'shape', 'table', 'text']);
      const table = tableOn(slide);
      expect(table.rows.slice(0, 2).map((row) => row.cells.map(({ text }) => text)))
        .toEqual([
          ['Header A', 'Header B'],
          ['Header 2 A', 'Header 2 B'],
        ]);
      expect(table.transform.x).toBe(inches(1.25));
      expect(table.transform.y).toBe(inches(index === 0 ? 3 : 1.25));
      expect(table.transform.width).toBe(inches(8));
      expect(activeImported.opcPackage.relationships(slide.partUri).some(
        ({ type, resolvedTarget }) =>
          type.endsWith('/slideLayout') && resolvedTarget === activeLayout.partUri,
      )).toBe(true);
    }

    const weightedPageCount = async (
      autoPageCharWeight: number,
      autoPageLineWeight: number,
    ) => {
      const generated = new PptxGenJS();
      await withTableToSlidesGlobals(tableToSlidesFixture({ bodyRows: 24 }), () =>
        generated.tableToSlides('report', {
          autoPageCharWeight,
          autoPageLineWeight,
          x: 0.5,
          y: 2.5,
          w: 4,
          h: 5,
          slideMargin: 0.5,
        }));
      return (await openPptxGenJSPublicOutput(generated)).slides.length;
    };
    const compactLines = await weightedPageCount(0, -1);
    const tallLines = await weightedPageCount(0, 1);
    const narrowCharacters = await weightedPageCount(-1, 0);
    const wideCharacters = await weightedPageCount(1, 0);
    expect(compactLines).toBeLessThan(tallLines);
    expect(narrowCharacters).not.toBe(wideCharacters);

    const alias = new PptxGenJS();
    await withTableToSlidesGlobals(tableToSlidesFixture({
      bodyRows: 45,
      headerRows: 2,
    }), () => alias.tableToSlides('report', {
      addHeaderToEach: true,
      newSlideStartY: 1.25,
      x: 1,
      y: 3,
      w: 8,
      h: 6,
    }));
    const aliasImported = await openPptxGenJSPublicOutput(alias);
    expect(aliasImported.slides.length).toBeGreaterThan(1);
    for (const [index, slide] of aliasImported.slides.entries()) {
      const table = tableOn(slide);
      expect(table.rows.slice(0, 2).map((row) => row.cells.map(({ text }) => text)))
        .toEqual([
          ['Header A', 'Header B'],
          ['Header 2 A', 'Header 2 B'],
        ]);
      expect(table.transform.y).toBe(inches(index === 0 ? 3 : 1.25));
    }

    const inertVariants: readonly (readonly [string, Record<string, unknown>])[] = [
      ['baseline', {}],
      ['align', { align: 'right' }],
      ['autoPage', { autoPage: false }],
      ['autoPageHeaderRows', { autoPageHeaderRows: 9 }],
      ['border scalar', { border: { color: 'FF0000', pt: 9 } }],
      ['border tuple', { border: Array.from({ length: 4 }, () => ({ color: '00FF00', pt: 8 })) }],
      ['colW scalar', { colW: 7 }],
      ['colW vector', { colW: [3, 4] }],
      ['margin scalar', { margin: 9 }],
      ['margin tuple', { margin: [9, 8, 7, 6] }],
      ['objectName', { objectName: 'Ignored table name' }],
      ['rowH scalar', { rowH: 7 }],
      ['rowH vector', { rowH: [3, 4, 5] }],
      ['transparency', { transparency: 75 }],
      ['valign', { valign: 'bottom' }],
    ];
    const inert = new PptxGenJS();
    const inertInputs: Record<string, unknown>[] = [];
    for (const [, options] of inertVariants) {
      const input = structuredClone(options);
      inertInputs.push(input);
      await withTableToSlidesGlobals(tableToSlidesFixture(), () =>
        inert.tableToSlides('report', input));
    }
    const inertImported = await openPptxGenJSPublicOutput(inert);
    expect(inertImported.slides).toHaveLength(inertVariants.length);
    const baselineState = tableState(tableOn(inertImported.slides[0]!));
    inertImported.slides.forEach((slide, index) => {
      expect(tableState(tableOn(slide)), inertVariants[index]![0]).toEqual(baselineState);
    });
    expect(inertImported.slides[10]!.shapes.map(({ name }) => name))
      .not.toContain('Ignored table name');
    expect(inertInputs[6]!.colW).not.toBe(7);
    expect(inertInputs[7]!.colW).not.toEqual([3, 4]);

    const forcedPaging = new PptxGenJS();
    await withTableToSlidesGlobals(tableToSlidesFixture({ bodyRows: 45 }), () =>
      forcedPaging.tableToSlides('report', { autoPage: false }));
    expect((await openPptxGenJSPublicOutput(forcedPaging)).slides.length)
      .toBeGreaterThan(1);

    const verboseLogs: unknown[][] = [];
    const verboseLog = vi.spyOn(console, 'log').mockImplementation((...values) => {
      verboseLogs.push(values);
    });
    try {
      const verbose = new PptxGenJS();
      await withTableToSlidesGlobals(tableToSlidesFixture(), () =>
        verbose.tableToSlides('report', { verbose: true }));
      const verboseImported = await openPptxGenJSPublicOutput(verbose);
      expect(tableState(tableOn(verboseImported.slides[0]!))).toEqual(baselineState);
      expect(verboseLogs.flat().some((value) => String(value).includes('VERBOSE MODE')))
        .toBe(true);
    } finally {
      verboseLog.mockRestore();
    }
  }, 120_000);

  it('matches the PptxGenJS horizontal alignment runtime catalog', async () => {
    const generated = new PptxGenJS();
    const generatedAlignments = Object.values(generated.AlignH);

    expect(Object.keys(generated.AlignH)).toEqual(TEXT_ALIGNMENTS);
    expect(generatedAlignments).toEqual(TEXT_ALIGNMENTS);
    expect(TEXT_ALIGNMENTS).toEqual(['left', 'center', 'right', 'justify']);
    expect(Object.isFrozen(TEXT_ALIGNMENTS)).toBe(true);

    const slide = generated.addSlide();
    generatedAlignments.forEach((alignment, index) => {
      slide.addText(alignment, {
        align: alignment,
        x: 0.5,
        y: 0.5 + index,
        w: 3,
        h: 0.5,
      });
    });
    const imported = await importPptxGenJS(generated);
    expect(imported.slides[0]?.shapes.map((shape) =>
      (shape as ShapeModel).richText[0]?.align)).toEqual(TEXT_ALIGNMENTS);
  });

  it('matches the PptxGenJS output type runtime catalog and return kinds', async () => {
    const generated = new PptxGenJS();
    generated.addSlide().addText('PptxGenJS output types', {
      x: 1,
      y: 1,
      w: 4,
      h: 1,
    });
    const native = PptxDocument.create();
    native.addSlide().addText('Native output types');

    expect(Object.keys(generated.OutputType)).toEqual(OUTPUT_TYPES);
    expect(Object.values(generated.OutputType)).toEqual(OUTPUT_TYPES);
    expect(OUTPUT_TYPES).toEqual([
      'arraybuffer',
      'base64',
      'binarystring',
      'blob',
      'nodebuffer',
      'uint8array',
    ]);
    expect(Object.isFrozen(OUTPUT_TYPES)).toBe(true);

    const outputKind = (value: unknown): string => {
      if (Buffer.isBuffer(value)) return 'nodebuffer';
      if (value instanceof ArrayBuffer) return 'arraybuffer';
      if (value instanceof Blob) return `blob:${value.type}`;
      if (value instanceof Uint8Array) return 'uint8array';
      return typeof value;
    };
    const decode = async (outputType: OutputType, value: unknown): Promise<Uint8Array> => {
      if (outputType === 'arraybuffer') return new Uint8Array(value as ArrayBuffer);
      if (outputType === 'base64') return Uint8Array.from(Buffer.from(value as string, 'base64'));
      if (outputType === 'binarystring') {
        return Uint8Array.from(value as string, (character) => character.charCodeAt(0));
      }
      if (outputType === 'blob') return new Uint8Array(await (value as Blob).arrayBuffer());
      return new Uint8Array(value as Uint8Array);
    };

    for (const outputType of OUTPUT_TYPES) {
      const generatedOutput = await generated.write({ outputType, compression: false });
      const nativeOutput = await native.write({ outputType });
      expect(outputKind(nativeOutput), outputType).toBe(outputKind(generatedOutput));
      const generatedBytes = await decode(outputType, generatedOutput);
      const nativeBytes = await decode(outputType, nativeOutput);
      expect((await PptxDocument.open(generatedBytes)).slides).toHaveLength(1);
      expect((await PptxDocument.open(nativeBytes)).slides).toHaveLength(1);
    }
  }, 30_000);

  it('records the public PptxGenJS stream result and provides a real Node readable', async () => {
    const generated = new PptxGenJS();
    generated.addSlide().addText('PptxGenJS stream', {
      x: 1,
      y: 1,
      w: 4,
      h: 1,
    });
    const native = PptxDocument.create();
    native.addSlide().addText('Native stream');

    const generatedStream = await generated.stream();
    expect(Buffer.isBuffer(generatedStream)).toBe(true);
    expect((await PptxDocument.open(generatedStream as Uint8Array)).slides).toHaveLength(1);

    const nativeNodeBuffer = await native.write({ outputType: 'nodebuffer' });
    expect(Buffer.isBuffer(nativeNodeBuffer)).toBe(true);
    expect((await PptxDocument.open(nativeNodeBuffer)).slides).toHaveLength(1);

    const nativeStream = await native.stream();
    expect(nativeStream).toBeInstanceOf(Readable);
    expect(Buffer.isBuffer(nativeStream)).toBe(false);
    const chunks: Uint8Array[] = [];
    for await (const chunk of nativeStream) chunks.push(new Uint8Array(chunk));
    const length = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0);
    const bytes = new Uint8Array(length);
    let offset = 0;
    for (const chunk of chunks) {
      bytes.set(chunk, offset);
      offset += chunk.byteLength;
    }
    expect((await PptxDocument.open(bytes)).slides).toHaveLength(1);
  }, 30_000);

  it('matches legal compression booleans without copying the explicit-output defect', async () => {
    const generated = new PptxGenJS();
    generated.addSlide().addText('PptxGenJS compression '.repeat(2_000), {
      x: 1,
      y: 1,
      w: 4,
      h: 1,
    });
    const generatedStreamDefault = new Uint8Array(await generated.stream() as Uint8Array);
    const generatedStreamStore = new Uint8Array(
      await generated.stream({ compression: false }) as Uint8Array,
    );
    const generatedStreamDeflate = new Uint8Array(
      await generated.stream({ compression: true }) as Uint8Array,
    );
    const generatedWriteStore = await generated.write({
      outputType: 'uint8array',
      compression: false,
    });
    const generatedWriteDeflate = await generated.write({
      outputType: 'uint8array',
      compression: true,
    });

    expect(generatedStreamDefault).toEqual(generatedStreamStore);
    expect(new Set(zipCompressionMethods(generatedStreamStore))).toEqual(new Set([0]));
    expect(new Set(zipCompressionMethods(generatedStreamDeflate))).toEqual(new Set([8]));
    expect(generatedStreamDeflate.byteLength).toBeLessThan(generatedStreamStore.byteLength);
    expect(generatedWriteDeflate).toEqual(generatedWriteStore);
    expect(new Set(zipCompressionMethods(generatedWriteDeflate))).toEqual(new Set([0]));

    const native = PptxDocument.create();
    native.addSlide().addText('Native compression');
    native.opcPackage.setPart(
      '/custom/compression.bin',
      new Uint8Array(65_536).fill(0x41),
      'application/octet-stream',
    );
    const nativeStore = await native.write({
      outputType: 'uint8array',
      compression: false,
    });
    const nativeDeflate = await native.write({
      outputType: 'uint8array',
      compression: true,
    });
    expect(new Set(zipCompressionMethods(nativeStore))).toEqual(new Set([0]));
    expect(new Set(zipCompressionMethods(nativeDeflate))).toEqual(new Set([8]));
    expect(nativeDeflate.byteLength).toBeLessThan(nativeStore.byteLength);
    expect((await PptxDocument.open(nativeStore)).slides).toHaveLength(1);
    expect((await PptxDocument.open(nativeDeflate)).slides).toHaveLength(1);
  }, 30_000);

  it('locks the presentation root, output, and theme declarations against PptxGenJS 4.0.1', async () => {
    const rootPropertyId = (property: string) =>
      `class:PptxGenJS@property:${property}`;
    const propertyId = (owner: string, property: string) =>
      `interface:${owner}@property:${property}`;
    const outputUnionId = (token: string) =>
      `union:WRITE_OUTPUT_TYPE#${token}`;
    const returnUnionId = (method: 'stream' | 'write', token: string) =>
      `union:class:PptxGenJS#${method}@path:return#${token}`;
    const atomGroups = {
      supported: [
        'class:PptxGenJS#addSlide',
        ...['author', 'company', 'revision', 'rtlMode', 'subject', 'theme', 'title']
          .map(rootPropertyId),
        ...['height', 'width'].map((property) => propertyId('PresLayout', property)),
        ...['bodyFontFace', 'headFontFace']
          .map((property) => propertyId('ThemeProps', property)),
        propertyId('WriteBaseProps', 'compression'),
        outputUnionId('JSZIP_OUTPUT_TYPE'),
        ...['ArrayBuffer', 'Blob', 'Uint8Array', 'string']
          .map((token) => returnUnionId('write', token)),
      ],
      deliberate: [
        ...[
          'addSection',
          'defineLayout',
          'defineSlideMaster',
          'stream',
          'write',
          'writeFile',
        ].map((method) => `class:PptxGenJS#${method}`),
        rootPropertyId('layout'),
        propertyId('PresLayout', 'name'),
        propertyId('WriteFileProps', 'compression'),
        propertyId('WriteFileProps', 'fileName'),
        propertyId('WriteProps', 'compression'),
        propertyId('WriteProps', 'outputType'),
        outputUnionId('STREAM'),
        returnUnionId('stream', 'Uint8Array'),
      ],
      defect: ['ArrayBuffer', 'Blob', 'string']
        .map((token) => returnUnionId('stream', token)),
    } as const;
    expect(Object.fromEntries(Object.entries(atomGroups).map(([status, ids]) => [
      status,
      ids.length,
    ]))).toEqual({ supported: 18, deliberate: 14, defect: 3 });
    expect(new Set(Object.values(atomGroups).flat()).size).toBe(35);
    expect(Object.values(atomGroups).flat().some((id) => id.includes('ChartType')))
      .toBe(false);

    const generated = new PptxGenJS();
    expect(generated.addSection({ title: 'Root audit' })).toBeUndefined();
    expect(generated.defineLayout({
      name: 'ROOT_AUDIT',
      width: 11.7,
      height: 8.3,
    })).toBeUndefined();
    generated.layout = 'ROOT_AUDIT';
    expect(generated.presLayout).toMatchObject({
      name: 'ROOT_AUDIT',
      width: inches(11.7),
      height: inches(8.3),
    });
    expect(generated.defineSlideMaster({
      title: 'ROOT_AUDIT_MASTER',
      objects: [],
    })).toBeUndefined();
    const generatedMasterSlide = generated.addSlide('ROOT_AUDIT_MASTER');
    const generatedObjectSlide = generated.addSlide({ masterName: 'ROOT_AUDIT_MASTER' });
    expect(generatedMasterSlide).not.toBe(generatedObjectSlide);

    const native = PptxDocument.create({
      slideSize: { width: inches(11.7), height: inches(8.3) },
    });
    const nativeSection = native.addSection({ title: 'Root audit' });
    expect(nativeSection.title).toBe('Root audit');
    expect(native.presLayout).toEqual({
      name: 'custom',
      width: inches(11.7),
      height: inches(8.3),
    });
    const nativeLayout = await native.defineSlideMaster({ title: 'ROOT_AUDIT_MASTER' });
    expect(nativeLayout.name).toBe('ROOT_AUDIT_MASTER');
    expect(native.addSlide({ masterName: nativeLayout.name }).presentation).toBe(native);

    const streamResults = await Promise.all([
      generated.stream(),
      generated.stream({ compression: false }),
      generated.stream({ compression: true }),
    ]);
    expect(streamResults.every((value) => Buffer.isBuffer(value))).toBe(true);
    expect(streamResults.every((value) => value instanceof Uint8Array)).toBe(true);
    expect(streamResults.some((value) =>
      value instanceof ArrayBuffer || value instanceof Blob || typeof value === 'string'))
      .toBe(false);

    const misleadingStream = await generated.write({
      outputType: 'STREAM',
      compression: true,
    });
    expect(Buffer.isBuffer(misleadingStream)).toBe(true);
    expect(misleadingStream).toBeInstanceOf(Uint8Array);
    expect((misleadingStream as unknown as { read?: unknown }).read).toBeUndefined();

    const nativeStream = await native.stream();
    expect(nativeStream).toBeInstanceOf(Readable);
    expect(Buffer.isBuffer(nativeStream)).toBe(false);
    const nativeStreamChunks: Uint8Array[] = [];
    for await (const chunk of nativeStream) nativeStreamChunks.push(new Uint8Array(chunk));
    expect(nativeStreamChunks.length).toBeGreaterThan(0);

    const generatedStored = await generated.write({
      outputType: 'uint8array',
      compression: false,
    });
    const generatedRequestedDeflate = await generated.write({
      outputType: 'uint8array',
      compression: true,
    });
    expect(new Set(zipCompressionMethods(generatedStored))).toEqual(new Set([0]));
    expect(new Set(zipCompressionMethods(generatedRequestedDeflate))).toEqual(new Set([0]));
    const nativeStored = await native.write({
      outputType: 'uint8array',
      compression: false,
    });
    const nativeDeflate = await native.write({
      outputType: 'uint8array',
      compression: true,
    });
    expect(new Set(zipCompressionMethods(nativeStored))).toEqual(new Set([0]));
    expect(new Set(zipCompressionMethods(nativeDeflate))).toEqual(new Set([8]));

    const directory = await mkdtemp(join(tmpdir(), 'pptx-root-output-'));
    try {
      const generatedPath = join(directory, 'generated.pptx');
      const nativePath = join(directory, 'native.pptx');
      expect(await generated.writeFile({ fileName: generatedPath, compression: true }))
        .toBe(generatedPath);
      expect(await native.writeFile(nativePath, { compression: true })).toBeUndefined();
      const [generatedFile, nativeFile] = await Promise.all([
        readFile(generatedPath),
        readFile(nativePath),
      ]);
      expect(new Set(zipCompressionMethods(generatedFile))).toEqual(new Set([0]));
      expect(new Set(zipCompressionMethods(nativeFile))).toEqual(new Set([8]));
      expect((await PptxDocument.open(generatedFile)).slides).toHaveLength(2);
      expect((await PptxDocument.open(nativeFile)).slides).toHaveLength(1);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  }, 60_000);

  it('locks the slide and section declarations against PptxGenJS 4.0.1', async () => {
    const propertyId = (owner: string, property: string) =>
      `interface:${owner}@property:${property}`;
    const slidePropertyId = (property: string) => `property:Slide#${property}`;
    const atomGroups = {
      supported: [
        ...['addChart', 'addImage', 'addNotes', 'addShape', 'addTable', 'addText']
          .map((property) => propertyId('PresSlide', property)),
        propertyId('PresSlide', 'hidden'),
        'method:Slide#addNotes',
        slidePropertyId('hidden'),
        propertyId('SectionProps', 'title'),
      ],
      deliberate: [
        propertyId('AddSlideProps', 'masterName'),
        propertyId('AddSlideProps', 'sectionTitle'),
        ...['addMedia', 'background', 'color', 'slideNumber']
          .map((property) => propertyId('PresSlide', property)),
        propertyId('SectionProps', 'order'),
        ...['addImage', 'addMedia', 'addShape', 'addText']
          .map((method) => `method:Slide#${method}`),
        ...['background', 'color', 'newAutoPagedSlides', 'slideNumber']
          .map(slidePropertyId),
      ],
      deprecated: [slidePropertyId('bkgd')],
    } as const;
    expect(Object.fromEntries(Object.entries(atomGroups).map(([status, ids]) => [
      status,
      ids.length,
    ]))).toEqual({ supported: 10, deliberate: 15, deprecated: 1 });
    expect(new Set(Object.values(atomGroups).flat()).size).toBe(26);
    expect(Object.values(atomGroups).flat().some((id) => id.includes('@path:')))
      .toBe(false);

    const generated = new PptxGenJS();
    generated.defineSlideMaster({ title: 'SLIDE-SECTION-AUDIT', objects: [] });
    generated.addSection({ title: 'Family & <One>' });
    const generatedSlide = generated.addSlide({
      masterName: 'SLIDE-SECTION-AUDIT',
      sectionTitle: 'Family & <One>',
    });
    const hyperlink: { url: string; _rId?: number } = {
      url: 'https://slide-family.example?a=1&b=2',
    };
    expect(generatedSlide.addImage({
      data: PNG_DATA_URI,
      x: 0,
      y: 0,
      w: 1,
      h: 1,
    })).toBe(generatedSlide);
    expect(generatedSlide.addMedia({
      type: 'audio',
      data: 'data:audio/mpeg;base64,AQIDBA==',
      x: 1,
      y: 0,
      w: 1,
      h: 1,
    })).toBe(generatedSlide);
    expect(generatedSlide.addNotes('Family notes')).toBe(generatedSlide);
    expect(generatedSlide.addShape(generated.ShapeType.rect!, {
      x: 2,
      y: 0,
      w: 1,
      h: 1,
    })).toBe(generatedSlide);
    expect(generatedSlide.addText('Family text', {
      x: 3,
      y: 0,
      w: 1,
      h: 1,
      hyperlink,
    })).toBe(generatedSlide);
    expect(hyperlink._rId).toBeTypeOf('number');

    generatedSlide.hidden = true;
    generatedSlide.color = 'FF3399';
    const generatedSlideNumber = { x: 1, y: 1, w: 1, h: 0.3 };
    generatedSlide.slideNumber = generatedSlideNumber;
    expect(generatedSlide.hidden).toBe(true);
    expect(generatedSlide.slideNumber).toBe(generatedSlideNumber);

    const generatedBackground: { data: string; path?: string } = {
      data: PNG_DATA_URI,
    };
    generatedSlide.background = generatedBackground;
    expect(generatedBackground.path).toBe('preencoded.png');
    expect(generatedSlide.background).toBe(generatedBackground);

    const aliasOnly = generated.addSlide() as PptxGenJSSlide & { bkgd?: string };
    aliasOnly.bkgd = 'AABBCC';
    expect(aliasOnly.background).toEqual({ color: 'AABBCC' });
    const canonicalWins = generated.addSlide() as PptxGenJSSlide & { bkgd?: string };
    canonicalWins.background = { color: '112233' };
    canonicalWins.bkgd = '445566';
    expect(canonicalWins.background).toEqual({ color: '112233' });
    expect(canonicalWins.bkgd).toBe('445566');

    const autoPageSource = generated.addSlide({ sectionTitle: 'Family & <One>' });
    const following = generated.addSlide({ sectionTitle: 'Family & <One>' });
    expect(autoPageSource.newAutoPagedSlides).toBeUndefined();
    const autoPageOptions: Record<string, unknown> = {
      x: 1,
      y: 4,
      w: 8,
      autoPage: true,
      autoPageRepeatHeader: true,
      autoPageHeaderRows: 1,
      autoPageSlideStartY: 1,
      fontSize: 18,
      margin: 0.05,
    };
    autoPageSource.addTable([
      ['Header'],
      ...Array.from({ length: 15 }, (_, index) => [`Body ${index}`]),
    ], autoPageOptions);
    expect(autoPageSource.newAutoPagedSlides.length).toBeGreaterThan(0);
    expect(autoPageSource.newAutoPagedSlides[0]).toBe(following);
    expect(autoPageOptions).toHaveProperty('_arrObjTabHeadRows');
    const continuation = autoPageSource.newAutoPagedSlides[0]!;
    expect(Object.fromEntries(
      ['addChart', 'addImage', 'addMedia', 'addNotes', 'addShape', 'addTable', 'addText']
        .map((property) => [property, typeof continuation[property as keyof PptxGenJSSlide]]),
    )).toEqual({
      addChart: 'function',
      addImage: 'function',
      addMedia: 'function',
      addNotes: 'function',
      addShape: 'function',
      addTable: 'function',
      addText: 'function',
    });

    const masterProjection = (generated as unknown as {
      masterSlide: Record<string, unknown>;
    }).masterSlide;
    expect(Object.fromEntries(
      ['addChart', 'addImage', 'addMedia', 'addNotes', 'addShape', 'addTable', 'addText']
        .map((property) => [property, masterProjection[property]]),
    )).toEqual({
      addChart: null,
      addImage: null,
      addMedia: null,
      addNotes: null,
      addShape: null,
      addTable: null,
      addText: null,
    });
    expect(['background', 'color', 'hidden', 'slideNumber'].map(
      (property) => masterProjection[property],
    )).toEqual([undefined, undefined, undefined, undefined]);

    const unknownMaster = new PptxGenJS();
    expect(() => unknownMaster.addSlide({ masterName: 'MISSING' })).not.toThrow();
    const unknownSection = new PptxGenJS();
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    try {
      expect(() => unknownSection.addSlide({ sectionTitle: 'MISSING' })).not.toThrow();
      expect(warning).toHaveBeenCalledTimes(1);
    } finally {
      warning.mockRestore();
    }

    const ordered = new PptxGenJS();
    ordered.addSection({ title: 'A' });
    ordered.addSection({ title: 'C' });
    ordered.addSection({ title: 'B', order: 1 });
    const orderZero = new PptxGenJS();
    orderZero.addSection({ title: 'A' });
    orderZero.addSection({ title: 'Zero', order: 0 });
    const sectionTitles = (presentation: PptxGenJSInstance) =>
      (presentation as unknown as { _sections: readonly { title: string }[] })
        ._sections.map(({ title }) => title);
    expect(sectionTitles(ordered)).toEqual(['A', 'B', 'C']);
    expect(sectionTitles(orderZero)).toEqual(['A', 'Zero']);

    const native = PptxDocument.create();
    await native.defineSlideMaster({ title: 'SLIDE-SECTION-AUDIT' });
    native.addSection({ title: 'Family & <One>' });
    const nativeSlide = native.addSlide({
      masterName: 'SLIDE-SECTION-AUDIT',
      sectionTitle: 'Family & <One>',
    });
    expect(nativeSlide.addNotes('Family notes')).toBe(nativeSlide);
    expect(await native.addImage(0, PNG_DATA_URI, {
      x: inches(0),
      y: inches(0),
      width: inches(1),
      height: inches(1),
    })).toBeInstanceOf(ImageModel);
    expect(await native.addAudio(0, 'data:audio/mpeg;base64,AQIDBA=='))
      .toBeInstanceOf(MediaModel);
    expect(nativeSlide.addShape('rect', {
      x: inches(2),
      y: inches(0),
      width: inches(1),
      height: inches(1),
    })).toBeInstanceOf(ShapeModel);
    expect(nativeSlide.addText('Family text', {
      x: inches(3),
      y: inches(0),
      width: inches(1),
      height: inches(1),
    })).toBeInstanceOf(ShapeModel);
    expect(Object.fromEntries(
      ['addChart', 'addImage', 'addNotes', 'addShape', 'addTable', 'addText']
        .map((property) => [property, typeof nativeSlide[property as keyof typeof nativeSlide]]),
    )).toEqual({
      addChart: 'function',
      addImage: 'function',
      addNotes: 'function',
      addShape: 'function',
      addTable: 'function',
      addText: 'function',
    });
    expect('addMedia' in nativeSlide).toBe(false);
    expect(Object.isFrozen(nativeSlide.newAutoPagedSlides)).toBe(true);

    nativeSlide.hidden = true;
    nativeSlide.color = { kind: 'srgb', value: 'FF3399' };
    const nativeSlideNumber = { x: inches(1), y: inches(1) };
    nativeSlide.slideNumber = nativeSlideNumber;
    expect(nativeSlide.hidden).toBe(true);
    expect(nativeSlide.color).toEqual({ kind: 'srgb', value: 'FF3399' });
    expect(nativeSlide.slideNumber).not.toBe(nativeSlideNumber);
    expect(nativeSlide.slideNumber).toMatchObject(nativeSlideNumber);

    const nativeUnknownMaster = PptxDocument.create();
    const unknownMasterBefore = packageState(nativeUnknownMaster);
    expect(() => nativeUnknownMaster.addSlide({ masterName: 'MISSING' }))
      .toThrow(RangeError);
    expect(packageState(nativeUnknownMaster)).toEqual(unknownMasterBefore);
    const nativeUnknownSection = PptxDocument.create();
    const unknownSectionBefore = packageState(nativeUnknownSection);
    expect(() => nativeUnknownSection.addSlide({ sectionTitle: 'MISSING' }))
      .toThrow(RangeError);
    expect(packageState(nativeUnknownSection)).toEqual(unknownSectionBefore);

    const nativeOrdered = PptxDocument.create();
    nativeOrdered.addSection({ title: 'A' });
    nativeOrdered.addSection({ title: 'C' });
    nativeOrdered.addSection({ title: 'B', order: 1 });
    const nativeOrderZero = PptxDocument.create();
    nativeOrderZero.addSection({ title: 'A' });
    nativeOrderZero.addSection({ title: 'Zero', order: 0 });
    expect(nativeOrdered.sections?.map(({ title }) => title)).toEqual(['A', 'B', 'C']);
    expect(nativeOrderZero.sections?.map(({ title }) => title)).toEqual(['Zero', 'A']);
  }, 60_000);

  it('locks master, background, and slide-number declarations against PptxGenJS 4.0.1', async () => {
    const propertyId = (owner: string, property: string) =>
      `interface:${owner}@property:${property}`;
    const inlineObjectId = (property: string) =>
      `inline:interface:SlideMasterProps@property:objects@property:objects.${property}`;
    const unionId = (owner: string, property: string, token: string) =>
      `union:${propertyId(owner, property)}#${token}`;
    const atomGroups = {
      supported: [
        inlineObjectId('placeholder.text'),
        propertyId('SlideMasterProps', 'title'),
        propertyId('SlideNumberProps', 'margin'),
        propertyId('SlideNumberProps', 'valign'),
      ],
      deliberate: [
        ...[
          'image',
          'line',
          'placeholder',
          'placeholder.options',
          'rect',
          'text',
        ].map(inlineObjectId),
        ...['background', 'margin', 'objects', 'slideNumber']
          .map((property) => propertyId('SlideMasterProps', property)),
        ...['color', 'data', 'path', 'transparency', 'type']
          .map((property) => propertyId('BackgroundProps', property)),
        unionId('BackgroundProps', 'type', 'none'),
        unionId('BackgroundProps', 'type', 'solid'),
        ...['align', 'h', 'w', 'x', 'y']
          .map((property) => propertyId('SlideNumberProps', property)),
      ],
      deprecated: [
        propertyId('SlideMasterProps', 'bkgd'),
        unionId('SlideMasterProps', 'bkgd', 'string'),
        propertyId('BackgroundProps', 'alpha'),
        propertyId('BackgroundProps', 'fill'),
      ],
      defect: [
        inlineObjectId('chart'),
        unionId('SlideMasterProps', 'bkgd', 'BackgroundProps'),
        propertyId('BackgroundProps', 'src'),
        propertyId('SlideNumberProps', 'transparency'),
      ],
    } as const;
    expect(Object.fromEntries(Object.entries(atomGroups).map(([status, ids]) => [
      status,
      ids.length,
    ]))).toEqual({ supported: 4, deliberate: 22, deprecated: 4, defect: 4 });
    const atomIds = Object.values(atomGroups).flat();
    expect(atomIds).toHaveLength(34);
    expect(new Set(atomIds).size).toBe(34);
    expect(atomIds.some((id) => /(?:PresSlide|property:Slide#)/u.test(id))).toBe(false);
    expect(atomIds.some((id) => id.includes('ChartType'))).toBe(false);

    const directory = await mkdtemp(join(tmpdir(), 'pptx-master-background-family-'));
    try {
      const pngPath = join(directory, 'background.png');
      await writeFile(pngPath, Buffer.from(PNG_DATA_URI.split(',')[1]!, 'base64'));

      const masterSource = new PptxGenJS();
      const canonicalImageDefinition = {
        title: 'MASTER-CANONICAL-IMAGE',
        background: { data: PNG_DATA_URI },
        objects: [],
      };
      const fullDefinition = {
        title: 'MASTER-FULL',
        margin: [1, 2, 3, 4] as const,
        slideNumber: {
          x: 0,
          y: 0,
          w: 1,
          h: 0.3,
          align: 'right' as const,
        },
        objects: [
          { image: { objectName: 'family-image', data: PNG_DATA_URI, x: 0, y: 0, w: 1, h: 1 } },
          { line: { objectName: 'family-line', x: 1, y: 0, w: 1, h: 1 } },
          { rect: { objectName: 'family-rect', x: 2, y: 0, w: 1, h: 1 } },
          {
            text: {
              text: 'Family text',
              options: { objectName: 'family-text', x: 3, y: 0, w: 1, h: 1 },
            },
          },
          {
            placeholder: {
              text: 'Family prompt',
              options: {
                name: 'family_prompt',
                objectName: 'family-prompt',
                type: 'body' as const,
                x: 4,
                y: 0,
                w: 1,
                h: 1,
              },
            },
          },
          {
            placeholder: {
              options: {
                name: 'family_empty',
                objectName: 'family-empty',
                type: 'title' as const,
                x: 5,
                y: 0,
                w: 1,
                h: 1,
              },
            },
          },
        ] satisfies readonly PptxGenJSMasterObject[],
      };
      const masterDefinitions = [
        { title: 'MASTER-CANONICAL-COLOR', background: { color: 'FF0000' }, objects: [] },
        canonicalImageDefinition,
        { title: 'MASTER-LEGACY-STRING', bkgd: '00FF00', objects: [] },
        { title: 'MASTER-LEGACY-COLOR-OBJECT', bkgd: { color: '0000FF' }, objects: [] },
        { title: 'MASTER-LEGACY-DATA-OBJECT', bkgd: { data: PNG_DATA_URI }, objects: [] },
        { title: 'MASTER-LEGACY-PATH-OBJECT', bkgd: { path: pngPath }, objects: [] },
        { title: 'MASTER-LEGACY-SRC-OBJECT', bkgd: { src: PNG_DATA_URI }, objects: [] },
        fullDefinition,
      ];
      const masterSnapshots = masterDefinitions.map((definition) => JSON.stringify(definition));
      for (const definition of masterDefinitions) {
        expect(masterSource.defineSlideMaster(definition)).toBeUndefined();
        masterSource.addSlide({ masterName: definition.title });
      }
      const importedMasters = await openPptxGenJSPublicOutput(masterSource);
      expect(masterDefinitions.map((definition) => JSON.stringify(definition)))
        .toEqual(masterSnapshots);
      expect(canonicalImageDefinition.background).not.toHaveProperty('path');
      const importedLayout = (name: string) => importedMasters.layouts.find(
        ({ name: candidate }) => candidate === name,
      )!;
      expect(importedLayout('MASTER-CANONICAL-COLOR').background).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
      });
      expect(importedLayout('MASTER-CANONICAL-IMAGE').background?.kind).toBe('image');
      expect(importedLayout('MASTER-LEGACY-STRING').background).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: '00FF00' },
      });
      for (const name of [
        'MASTER-LEGACY-COLOR-OBJECT',
        'MASTER-LEGACY-DATA-OBJECT',
        'MASTER-LEGACY-PATH-OBJECT',
        'MASTER-LEGACY-SRC-OBJECT',
      ]) {
        expect(importedLayout(name).background, name).toBeUndefined();
        expect(importedMasters.opcPackage.relationships(importedLayout(name).partUri).some(
          ({ type }) => type === IMAGE_RELATIONSHIP,
        ), name).toBe(false);
      }
      const fullLayout = importedLayout('MASTER-FULL');
      expect(fullLayout.shapes.map(({ name }) => name)).toEqual(expect.arrayContaining([
        'family-image',
        'family-line',
        'family-rect',
        'family-text',
        'family-prompt',
        'family-empty',
      ]));
      expect(fullLayout.placeholders.map(({ text }) => text)).toEqual([
        'Family prompt',
        '',
      ]);
      expect(fullLayout.slideNumber?.align).toBe('right');

      const declaredChart = new PptxGenJS();
      expect(() => declaredChart.defineSlideMaster({
        title: 'MASTER-DECLARED-CHART',
        objects: [{
          chart: { x: 1, y: 1, w: 3, h: 2 } as unknown as PptxGenJSMasterChart,
        }],
      })).toThrow(TypeError);

      const nativeMaster = PptxDocument.create();
      const nativeMasterInput = {
        title: 'NATIVE-MASTER-FAMILY',
        background: {
          kind: 'solid' as const,
          color: { kind: 'srgb' as const, value: 'FF0000' },
        },
        margin: [inches(0.1), inches(0.2), inches(0.3), inches(0.4)] as const,
        slideNumber: { x: 0, y: 0, width: 100, height: 30, align: 'right' as const },
        objects: [
          {
            kind: 'rect' as const,
            options: {
              name: 'native-family-rect',
              x: inches(0),
              y: inches(0),
              width: inches(1),
              height: inches(1),
            },
          },
          {
            kind: 'placeholder' as const,
            text: 'Native prompt',
            options: { name: 'native_family_prompt', type: 'body' as const },
          },
        ],
      };
      const nativeMasterSnapshot = JSON.stringify(nativeMasterInput);
      const nativeLayout = await nativeMaster.defineSlideMaster(nativeMasterInput);
      expect(nativeLayout).toBe(nativeMaster.layouts.find(
        ({ partUri }) => partUri === nativeLayout.partUri,
      ));
      expect(nativeLayout.margin).toEqual({
        top: inches(0.1),
        right: inches(0.2),
        bottom: inches(0.3),
        left: inches(0.4),
      });
      expect(nativeLayout.background).toEqual(nativeMasterInput.background);
      expect(nativeLayout.slideNumber).toMatchObject({ align: 'right', width: 100, height: 30 });
      expect(JSON.stringify(nativeMasterInput)).toBe(nativeMasterSnapshot);

      const backgroundSource = new PptxGenJS();
      const fillBackground: { fill: string; color?: string } = { fill: '00FF00' };
      const dataBackground: { data: string; path?: string } = { data: PNG_DATA_URI };
      const backgroundValues = [
        undefined,
        { color: 'FF3399' },
        { color: 'FF3399', transparency: 50 },
        { color: '112233', alpha: 40 },
        { color: '223344', transparency: 25, alpha: 40 },
        fillBackground,
        dataBackground,
        { path: pngPath },
        { src: PNG_DATA_URI },
        { type: 'none' as const },
        { type: 'none' as const, color: '334455' },
        { type: 'solid' as const, color: '445566' },
      ];
      for (const background of backgroundValues) {
        const slide = backgroundSource.addSlide();
        if (background !== undefined) {
          slide.background = background;
          expect(slide.background).toBe(background);
        }
      }
      const importedBackgrounds = await openPptxGenJSPublicOutput(backgroundSource);
      expect(fillBackground.color).toBe('00FF00');
      expect(dataBackground.path).toBe('preencoded.png');
      expect(importedBackgrounds.slides.map(({ background }) => background?.kind))
        .toEqual([
          undefined,
          'solid',
          'solid',
          'solid',
          undefined,
          'solid',
          'image',
          'image',
          undefined,
          undefined,
          undefined,
          'solid',
        ]);
      expect(slideXml(importedBackgrounds, 3)).toContain('<a:alpha val="60000"/>');
      expect(slideXml(importedBackgrounds, 4).match(/<a:alpha\b/gu)).toHaveLength(2);
      expect(slideBackgroundStructuralState(importedBackgrounds, 8)).toMatchObject({
        kind: undefined,
        direct: { present: false },
      });
      expect(slideBackgroundStructuralState(importedBackgrounds, 9)).toMatchObject({
        kind: undefined,
        direct: { present: false, noFill: false },
      });
      expect(slideBackgroundStructuralState(importedBackgrounds, 10)).toMatchObject({
        kind: undefined,
        direct: { present: true, noFill: false, solidFill: false },
      });

      const nativeBackgrounds = PptxDocument.create();
      const noFillSlide = nativeBackgrounds.addSlide();
      noFillSlide.background = { kind: 'none' };
      const nativeSolidInput = {
        kind: 'solid' as const,
        color: { kind: 'srgb' as const, value: 'FF3399' },
        transparency: 50,
      };
      const solidSlide = nativeBackgrounds.addSlide();
      solidSlide.background = nativeSolidInput;
      const imageSlide = nativeBackgrounds.addSlide();
      await nativeBackgrounds.setSlideBackgroundImage(
        nativeBackgrounds.slides.indexOf(imageSlide),
        PNG_DATA_URI,
      );
      expect(noFillSlide.background).toEqual({ kind: 'none' });
      expect(slideBackgroundStructuralState(nativeBackgrounds, 0).direct.noFill).toBe(true);
      expect(solidSlide.background).not.toBe(nativeSolidInput);
      expect(solidSlide.background).toEqual(nativeSolidInput);
      expect(Object.isFrozen(solidSlide.background)).toBe(true);
      expect(imageSlide.background?.kind).toBe('image');

      const slideNumberSource = new PptxGenJS();
      const slideNumberValues: PptxGenJSSlideNumberProps[] = [
        { x: 0, y: 0, w: 0, h: 0 },
        { align: 'left' },
        { align: 'center' },
        { align: 'right' },
        { align: 'justify' },
        { margin: 0 },
        { margin: 7 },
        { margin: [1, 2, 3, 4] },
        { valign: 'top' },
        { valign: 'middle' },
        { valign: 'bottom' },
        { transparency: 25, color: 'FF3399' },
      ];
      for (const value of slideNumberValues) {
        const slide = slideNumberSource.addSlide();
        slide.slideNumber = value;
        expect(slide.slideNumber).toBe(value);
      }
      const importedSlideNumbers = await openPptxGenJSPublicOutput(slideNumberSource);
      expect(importedSlideNumbers.slides[0]?.slideNumber).toMatchObject({
        x: 0,
        y: 0,
        width: 800_000,
        height: 300_000,
      });
      expect(importedSlideNumbers.slides.slice(1, 5)
        .map(({ slideNumber }) => slideNumber?.align))
        .toEqual(['left', 'center', 'right', 'left']);
      expect(importedSlideNumbers.slides.slice(5, 8)
        .map(({ slideNumber }) => slideNumber?.margin))
        .toEqual([
          { top: 0, right: 0, bottom: 0, left: 0 },
          { top: 7, right: 7, bottom: 7, left: 7 },
          { top: 1, right: 2, bottom: 3, left: 4 },
        ]);
      expect(slideNumberOwnerState(
        importedSlideNumbers,
        importedSlideNumbers.slides[5]!.partUri,
      ).xml).toContain(' lIns="0" tIns="0" rIns="0" bIns="0"');
      expect(importedSlideNumbers.slides.slice(8, 11)
        .map(({ slideNumber }) => slideNumber?.valign))
        .toEqual(['top', 'middle', 'bottom']);
      expect(importedSlideNumbers.slides[11]?.slideNumber?.style.transparency)
        .toBeUndefined();
      expect(slideNumberOwnerState(
        importedSlideNumbers,
        importedSlideNumbers.slides[11]!.partUri,
      ).xml).not.toContain('<a:alpha');

      const nativeSlideNumbers = PptxDocument.create();
      const nativeSlideNumber = nativeSlideNumbers.addSlide();
      const nativeSlideNumberInput = {
        x: 0,
        y: 0,
        width: 100,
        height: 30,
        align: 'justify' as const,
        valign: 'middle' as const,
        margin: 0,
        style: { transparency: 25 },
      };
      nativeSlideNumber.slideNumber = nativeSlideNumberInput;
      expect(nativeSlideNumber.slideNumber).not.toBe(nativeSlideNumberInput);
      expect(nativeSlideNumber.slideNumber).toMatchObject({
        x: 0,
        y: 0,
        width: 100,
        height: 30,
        align: 'justify',
        valign: 'middle',
        margin: { top: 0, right: 0, bottom: 0, left: 0 },
        style: { transparency: 25 },
      });
      expect(Object.isFrozen(nativeSlideNumber.slideNumber)).toBe(true);
      const reopenedNativeSlideNumbers = await PptxDocument.open(
        await nativeSlideNumbers.write(),
      );
      expect(reopenedNativeSlideNumbers.slides[0]?.slideNumber).toMatchObject({
        align: 'justify',
        margin: { top: 0, right: 0, bottom: 0, left: 0 },
        style: { transparency: 25 },
      });
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  }, 60_000);

  it('matches the public PptxGenJS SchemeColor helper and legal output', async () => {
    const generated = new PptxGenJS();
    const second = new PptxGenJS();
    expect(Object.entries(generated.SchemeColor)).toEqual(Object.entries(SCHEME_COLORS));
    expect(second.SchemeColor).toBe(generated.SchemeColor);
    expect(Object.isFrozen(generated.SchemeColor)).toBe(false);
    expect(Object.getOwnPropertyDescriptor(
      Object.getPrototypeOf(generated),
      'SchemeColor',
    )).toMatchObject({ set: undefined, enumerable: false });
    expect(Object.isFrozen(SCHEME_COLORS)).toBe(true);

    const slide = generated.addSlide();
    slide.addText([
      { text: 'Text1', options: { color: generated.SchemeColor.text1 } },
      { text: 'Accent1', options: { color: generated.SchemeColor.accent1 } },
    ], { x: 1, y: 1, w: 5, h: 1 });
    const imported = await importPptxGenJS(generated);
    const shape = imported.slides[0]?.shapes[0];
    expect(shape).toBeInstanceOf(ShapeModel);
    expect((shape as ShapeModel).richText[0]?.runs.map(({ style }) => style?.color)).toEqual([
      { kind: 'scheme', value: SCHEME_COLORS.text1 },
      { kind: 'scheme', value: SCHEME_COLORS.accent1 },
    ]);
  });

  it('matches the PptxGenJS vertical alignment runtime catalog', async () => {
    const generated = new PptxGenJS();
    generated.layout = 'LAYOUT_WIDE';
    const generatedAlignments = Object.values(generated.AlignV);

    expect(Object.keys(generated.AlignV)).toEqual(TEXT_VERTICAL_ALIGNMENTS);
    expect(generatedAlignments).toEqual(TEXT_VERTICAL_ALIGNMENTS);
    expect(TEXT_VERTICAL_ALIGNMENTS).toEqual(['top', 'middle', 'bottom']);
    expect(Object.isFrozen(TEXT_VERTICAL_ALIGNMENTS)).toBe(true);

    const textSlide = generated.addSlide();
    generatedAlignments.forEach((alignment, index) => {
      textSlide.addText(alignment, {
        valign: alignment,
        x: 0.5,
        y: 0.5 + index,
        w: 3,
        h: 0.5,
      });
    });
    const tableSlide = generated.addSlide();
    tableSlide.addTable([
      generatedAlignments.map((alignment) => ({
        text: alignment,
        options: { valign: alignment },
      })),
    ], { x: 0.5, y: 4, w: 12, h: 1, valign: 'bottom' });

    const imported = await importPptxGenJS(generated);
    expect(imported.slides[0]?.shapes.slice(0, 3).map((shape) =>
      (shape as ShapeModel).verticalAlignment)).toEqual(TEXT_VERTICAL_ALIGNMENTS);
    const importedTable = imported.slides[1]?.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(importedTable?.rows[0]?.cells.map(
      ({ verticalAlignment }) => verticalAlignment,
    )).toEqual(TEXT_VERTICAL_ALIGNMENTS);
  });

  it('matches the public presentation layout projection and locks the custom-name boundary', async () => {
    const publicLayout = ({ name, width, height }: PptxGenJSInstance['presLayout']) => ({
      name,
      width,
      height,
    });
    const defaultGenerated = new PptxGenJS();
    const defaultNative = PptxDocument.create();
    expect(publicLayout(defaultGenerated.presLayout)).toEqual(defaultNative.presLayout);

    const cases = [
      ['LAYOUT_4x3', '4:3', 'screen4x3', 9_144_000, 6_858_000],
      ['LAYOUT_16x9', '16:9', 'screen16x9', 9_144_000, 5_143_500],
      ['LAYOUT_16x10', '16:10', 'screen16x10', 9_144_000, 5_715_000],
      ['LAYOUT_WIDE', 'wide', 'custom', 12_192_000, 6_858_000],
    ] as const;
    for (const [generatedLayout, nativeSize, name, width, height] of cases) {
      const generated = new PptxGenJS();
      generated.layout = generatedLayout;
      const native = PptxDocument.create({ slideSize: nativeSize });
      const expected = { name, width, height };
      expect(publicLayout(generated.presLayout)).toEqual(expected);
      expect(native.presLayout).toEqual(expected);
      expect(Object.keys(native.presLayout)).toEqual(['name', 'width', 'height']);
      expect(native.presLayout).not.toBe(native.presLayout);

      await generated.write({ outputType: 'uint8array', compression: false });
      await native.write();
      expect(publicLayout(generated.presLayout)).toEqual(expected);
      expect(native.presLayout).toEqual(expected);
    }

    const customGenerated = new PptxGenJS();
    customGenerated.defineLayout({ name: 'CUSTOM', width: 11.7, height: 8.3 });
    customGenerated.layout = 'CUSTOM';
    const customNative = PptxDocument.create({
      slideSize: { width: inches(11.7), height: inches(8.3) },
    });
    const expectedGeneratedCustom = {
      name: 'CUSTOM',
      width: inches(11.7),
      height: inches(8.3),
    };
    const expectedNativeCustom = {
      name: 'custom',
      width: inches(11.7),
      height: inches(8.3),
    } as const;
    expect(publicLayout(customGenerated.presLayout)).toEqual(expectedGeneratedCustom);
    expect(customNative.presLayout).toEqual(expectedNativeCustom);
    const customGeneratedBytes = await customGenerated.write({
      outputType: 'nodebuffer',
      compression: false,
    });
    const customNativeBytes = await customNative.write();
    expect(publicLayout(customGenerated.presLayout)).toEqual(expectedGeneratedCustom);
    expect(customNative.presLayout).toEqual(expectedNativeCustom);
    expect((await PptxDocument.open(customGeneratedBytes)).presLayout)
      .toEqual(expectedNativeCustom);
    expect((await PptxDocument.open(customNativeBytes)).presLayout)
      .toEqual(expectedNativeCustom);
    expect(Object.getOwnPropertyDescriptor(
      Object.getPrototypeOf(customGenerated),
      'presLayout',
    )).toMatchObject({ set: undefined, enumerable: false });
    expect(Object.getOwnPropertyDescriptor(PptxDocument.prototype, 'presLayout')).toMatchObject({
      set: undefined,
      enumerable: false,
    });
  });

  it('matches public PptxGenJS slide default colors and locks intentional differences', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');

    const srgb = generated.addSlide();
    srgb.color = 'ff3399';
    srgb.addText('sRGB inherited', { x: 1, y: 1, w: 5, h: 0.5 });

    const scheme = generated.addSlide();
    scheme.color = generated.SchemeColor.accent1;
    scheme.addText('Scheme inherited', { x: 1, y: 1, w: 5, h: 0.5 });

    const rich = generated.addSlide();
    rich.color = generated.SchemeColor.accent1;
    rich.addText([
      { text: 'Rich inherited' },
      { text: 'Rich override', options: { color: '00aa00' } },
      { text: 'Transparent inherited', options: { transparency: 25 } },
    ], { x: 1, y: 1, w: 5, h: 0.5 });

    const temporal = generated.addSlide();
    temporal.color = 'ff3399';
    temporal.addText('First default', { x: 1, y: 1, w: 5, h: 0.5 });
    temporal.color = generated.SchemeColor.accent2;
    temporal.addText('Second default', { x: 1, y: 2, w: 5, h: 0.5 });
    temporal.color = undefined;
    temporal.addText('Cleared default', { x: 1, y: 3, w: 5, h: 0.5 });

    const cleared = generated.addSlide();
    cleared.color = 'ff3399';
    cleared.color = undefined;
    cleared.addText('Cleared control', { x: 1, y: 1, w: 5, h: 0.5 });

    const tableSlide = generated.addSlide();
    tableSlide.color = generated.SchemeColor.accent2;
    tableSlide.addTable([[{ text: 'Table default', options: {} }]], {
      x: 1,
      y: 1,
      w: 5,
      h: 1,
    });

    const imported = await importPptxGenJS(generated);
    const native = PptxDocument.create();
    const nativeSrgb = native.addSlide();
    nativeSrgb.color = { kind: 'srgb', value: 'ff3399' };
    nativeSrgb.addText('sRGB inherited');
    const nativeScheme = native.addSlide();
    nativeScheme.color = { kind: 'scheme', value: 'accent1' };
    nativeScheme.addText('Scheme inherited');
    const nativeRich = native.addSlide();
    nativeRich.color = { kind: 'scheme', value: 'accent1' };
    nativeRich.addRichText([{
      runs: [
        { text: 'Rich inherited' },
        { text: 'Rich override', style: { color: { kind: 'srgb', value: '00AA00' } } },
        { text: 'Transparent inherited', style: { transparency: 25 } },
      ],
    }]);
    const nativeTemporal = native.addSlide();
    nativeTemporal.color = { kind: 'srgb', value: 'FF3399' };
    nativeTemporal.addText('First default');
    nativeTemporal.color = { kind: 'scheme', value: 'accent2' };
    nativeTemporal.addText('Second default');
    nativeTemporal.color = undefined;
    nativeTemporal.addText('Cleared default');
    const nativeCleared = native.addSlide();
    nativeCleared.color = { kind: 'srgb', value: 'FF3399' };
    nativeCleared.color = undefined;
    nativeCleared.addText('Cleared control');
    const nativeTableSlide = native.addSlide();
    nativeTableSlide.color = { kind: 'scheme', value: 'accent2' };
    nativeTableSlide.addTable([['Table default']]);
    const reopenedNative = await PptxDocument.open(await native.write());

    const runState = (document: PptxDocument, slideIndex: number) => document.slides[slideIndex]!
      .shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .map(({ richText }) => richText.flatMap(({ runs }) => runs.map(({ style }) => ({
        color: style?.color,
        transparency: style?.transparency,
      }))));
    const customExpected = [
      [[{ color: { kind: 'srgb', value: 'FF3399' }, transparency: undefined }]],
      [[{ color: { kind: 'scheme', value: 'accent1' }, transparency: undefined }]],
      [[
        { color: { kind: 'scheme', value: 'accent1' }, transparency: undefined },
        { color: { kind: 'srgb', value: '00AA00' }, transparency: undefined },
        { color: { kind: 'scheme', value: 'accent1' }, transparency: 25 },
      ]],
    ];
    expect([0, 1, 2].map((index) => runState(imported, index))).toEqual(customExpected);
    expect([0, 1, 2].map((index) => runState(reopenedNative, index))).toEqual(customExpected);
    expect(runState(imported, 3)).toEqual([
      [{ color: { kind: 'srgb', value: 'FF3399' }, transparency: undefined }],
      [{ color: { kind: 'scheme', value: 'accent2' }, transparency: undefined }],
      [{ color: { kind: 'srgb', value: '000000' }, transparency: undefined }],
    ]);
    expect(runState(reopenedNative, 3)).toEqual([
      [{ color: { kind: 'srgb', value: 'FF3399' }, transparency: undefined }],
      [{ color: { kind: 'scheme', value: 'accent2' }, transparency: undefined }],
      [{ color: { kind: 'scheme', value: 'tx1' }, transparency: undefined }],
    ]);
    expect(runState(imported, 4)[0]?.[0]?.color)
      .toEqual({ kind: 'srgb', value: '000000' });
    expect(runState(reopenedNative, 4)[0]?.[0]?.color)
      .toEqual({ kind: 'scheme', value: 'tx1' });

    const importedTable = imported.slides[5]!.shapes[0] as TableModel;
    const nativeTable = reopenedNative.slides[5]!.shapes[0] as TableModel;
    const importedTableXml = imported.slides[5]!.resolveShape(importedTable.id);
    const nativeTableXml = reopenedNative.slides[5]!.resolveShape(nativeTable.id);
    expect(importedTableXml.xml.original(importedTableXml.element))
      .toContain('<a:srgbClr val="000000"/>');
    expect(nativeTableXml.xml.original(nativeTableXml.element))
      .toContain('<a:schemeClr val="tx1"/>');
    for (const { xml, element } of [importedTableXml, nativeTableXml]) {
      expect(xml.original(element)).not.toContain('<a:schemeClr val="accent2"/>');
    }

    for (const compatibility of ['powerpoint-2010', 'powerpoint-current'] as const) {
      await imported.write({ compatibility });
      await native.write({ compatibility });
      expect(imported.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
      expect(native.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
    }
  });

  it('contrasts PptxGenJS slide default color fallback with strict native rejection', async () => {
    const generated = new PptxGenJS();
    const generatedSlide = generated.addSlide();
    generatedSlide.color = 'BAD';
    generatedSlide.addText('Invalid fallback', { x: 1, y: 1, w: 5, h: 0.5 });
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    try {
      const imported = await importPptxGenJS(generated);
      expect(warning).toHaveBeenCalled();
      expect((imported.slides[0]!.shapes[0] as ShapeModel).richText[0]?.runs[0]?.style?.color)
        .toEqual({ kind: 'srgb', value: '000000' });

      warning.mockClear();
      const native = PptxDocument.create();
      const nativeSlide = native.addSlide();
      nativeSlide.color = { kind: 'scheme', value: 'accent1' };
      const beforeBytes = native.opcPackage.requirePart(nativeSlide.partUri).bytes;
      const beforeJournal = [...native.opcPackage.mutations];
      expect(() => {
        nativeSlide.color = 'BAD' as never;
      }).toThrow(TypeError);
      expect(warning).not.toHaveBeenCalled();
      expect(nativeSlide.color).toEqual({ kind: 'scheme', value: 'accent1' });
      expect(nativeSlide.shapes).toHaveLength(0);
      expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeBytes);
      expect(native.opcPackage.mutations).toEqual(beforeJournal);
    } finally {
      warning.mockRestore();
    }
  });

  it('imports public slide-number variants and locks PptxGenJS 4.0.1 differences', async () => {
    const source = new PptxGenJS();
    source.layout = 'LAYOUT_WIDE';
    source.rtlMode = true;
    source.addSlide();
    const options: readonly PptxGenJSSlideNumberProps[] = [
      {},
      {
        x: 0,
        y: 0,
        w: 1,
        h: 0.3,
        align: 'center',
        fontFace: 'Aptos',
        fontSize: 18,
        lang: 'zh-CN',
        bold: true,
        italic: true,
        color: 'FF3399',
        transparency: 25,
        margin: [1, 2, 3, 4],
      },
      { color: source.SchemeColor.accent1 },
      { align: 'left' },
      { align: 'center' },
      { align: 'right' },
      { margin: 7 },
      { margin: [1, 2, 3, 4] },
      { valign: 'top' },
      { valign: 'middle' },
      { valign: 'bottom' },
      { align: 'justify' },
      { w: 0, h: 0 },
      { italic: true, transparency: 25, color: 'FF3399', lang: 'zh-CN' },
    ];
    for (const value of options) source.addSlide().slideNumber = value;

    const imported = await importPptxGenJS(source);
    expect(imported.slides).toHaveLength(options.length + 1);
    expect(imported.slides[0]?.slideNumber).toBeUndefined();
    expect(imported.slides[1]?.slideNumber).toEqual({
      x: 0,
      y: 0,
      width: 800_000,
      height: 300_000,
      align: 'left',
      rtl: false,
      style: { lang: 'en-US', bold: false, italic: false },
    });
    expect(imported.slides[2]?.slideNumber).toMatchObject({
      x: 0,
      y: 0,
      width: 914_400,
      height: 274_320,
      align: 'center',
      rtl: false,
      margin: { top: 1, right: 2, bottom: 3, left: 4 },
      style: {
        fontFamily: 'Aptos',
        fontSize: 18,
        lang: 'en-US',
        bold: true,
        italic: false,
        color: { kind: 'srgb', value: 'FF3399' },
      },
    });
    expect(imported.slides[3]?.slideNumber?.style.color)
      .toEqual({ kind: 'scheme', value: 'accent1' });
    expect(imported.slides.slice(4, 7).map(({ slideNumber }) => slideNumber?.align))
      .toEqual(['left', 'center', 'right']);
    expect(imported.slides[7]?.slideNumber?.margin)
      .toEqual({ top: 7, right: 7, bottom: 7, left: 7 });
    expect(imported.slides[8]?.slideNumber?.margin)
      .toEqual({ top: 1, right: 2, bottom: 3, left: 4 });
    expect(imported.slides[9]?.slideNumber?.valign).toBe('top');
    expect(imported.slides[10]?.slideNumber?.valign).toBe('middle');
    expect(slideNumberOwnerState(imported, imported.slides[10]!.partUri).xml)
      .toContain(' anchor="ctr"');
    expect(imported.slides[11]?.slideNumber?.valign).toBe('bottom');
    expect(imported.slides[12]?.slideNumber?.align).toBe('left');
    expect(imported.slides[13]?.slideNumber).toMatchObject({
      width: 800_000,
      height: 300_000,
    });
    expect(imported.slides[14]?.slideNumber).toMatchObject({
      rtl: false,
      style: { lang: 'en-US', italic: false, color: { kind: 'srgb', value: 'FF3399' } },
    });
    expect(imported.slides[14]?.slideNumber?.style.transparency).toBeUndefined();

    for (const [index, slide] of imported.slides.entries()) {
      const state = slideNumberOwnerState(imported, slide.partUri);
      expect(state.ownerCount).toBe(index === 0 ? 0 : 1);
      if (index === 0) continue;
      expect(state).toMatchObject({
        shapeId: 25,
        shapeIdOccurrences: 1,
        fieldType: 'slidenum',
        cache: String(index + 1),
      });
    }
    const layoutState = slideNumberOwnerState(imported, imported.layouts[0]!.partUri);
    const masterState = slideNumberOwnerState(imported, imported.masters[0]!.partUri);
    expect(layoutState.cache).toMatch(/^10\d{2}$/);
    expect(layoutState).toMatchObject({ ownerCount: 1, shapeId: 25 });
    expect(masterState).toMatchObject({ ownerCount: 1, shapeId: 25, cache: 'null', masterFlag: '0' });
    expect(imported.layouts[0]?.slideNumber).toBeDefined();
    expect(imported.masters[0]?.slideNumber).toBeUndefined();

    const native = PptxDocument.create({ firstSlideNumber: 5 });
    const nativeSlide = native.addSlide();
    nativeSlide.slideNumber = {
      align: 'justify',
      rtl: true,
      valign: 'middle',
      style: {
        lang: 'zh-CN',
        italic: true,
        color: { kind: 'srgb', value: 'FF3399' },
        transparency: 25,
      },
    };
    expect(nativeSlide.slideNumber).toMatchObject({
      align: 'justify',
      rtl: true,
      valign: 'middle',
      style: { lang: 'zh-CN', italic: true, transparency: 25 },
    });
    expect(slideNumberOwnerState(native, nativeSlide.partUri)).toMatchObject({
      shapeId: 2,
      shapeIdOccurrences: 1,
      cache: '5',
    });
    expect(native.layouts[0]?.slideNumber).toBeUndefined();
    expect(native.masters[0]?.slideNumber).toBeUndefined();
    const nativeJournal = [...native.opcPackage.mutations];
    expect(() => {
      nativeSlide.slideNumber = { width: 0 };
    }).toThrow(/width/i);
    expect(native.opcPackage.mutations).toEqual(nativeJournal);
  });

  it('diagnoses, canonicalizes, reorders, and reopens named-master public output', async () => {
    const source = new PptxGenJS();
    source.defineSlideMaster({
      title: 'NUMBERED_MASTER',
      background: { color: 'FFFFFF' },
      objects: [],
      slideNumber: { x: 1, y: 1, w: 1, h: 0.3, align: 'right' },
    });
    source.addSlide({ masterName: 'NUMBERED_MASTER' });
    const imported = await importPptxGenJS(source);
    const numberedLayout = imported.layouts.find(({ slideNumber }) => slideNumber !== undefined)!;
    const master = imported.masters[0]!;
    expect(imported.slides[0]?.slideNumber).toMatchObject({
      x: 914_400,
      y: 914_400,
      width: 914_400,
      height: 274_320,
      align: 'right',
    });
    expect(numberedLayout.slideNumber?.align).toBe('right');
    expect(master.slideNumber).toBeUndefined();
    expect(slideNumberOwnerState(imported, numberedLayout.partUri).cache).toMatch(/^10\d{2}$/);
    expect(slideNumberOwnerState(imported, master.partUri))
      .toMatchObject({ ownerCount: 1, shapeId: 25, cache: 'null', masterFlag: '0' });

    const profiles = [
      'powerpoint-2010',
      'powerpoint-current',
      'keynote-current',
      'google-slides-import',
      'libreoffice-current',
    ] as const;
    for (const compatibility of profiles) {
      await imported.write({ mode: 'permissive', compatibility });
      expect(imported.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
      expect(imported.diagnostics.filter(({ code }) => code.startsWith('SLIDE_NUMBER_')))
        .toEqual(expect.arrayContaining([
          expect.objectContaining({
            code: 'SLIDE_NUMBER_CACHE_NONCANONICAL',
            partUri: numberedLayout.partUri,
            compatibility,
          }),
          expect.objectContaining({
            code: 'SLIDE_NUMBER_MASTER_DISABLED',
            partUri: master.partUri,
            compatibility,
          }),
          expect.objectContaining({
            code: 'SLIDE_NUMBER_CACHE_NONCANONICAL',
            partUri: master.partUri,
            compatibility,
          }),
        ]));
    }

    const themeUri = imported.themes[0]!.partUri;
    const themeBefore = imported.opcPackage.requirePart(themeUri).bytes;
    const slideRelationshipsBefore = imported.opcPackage.relationships(
      imported.slides[0]!.partUri,
    ).map((relationship) => ({ ...relationship }));
    imported.slides[0]!.slideNumber = { align: 'center', style: { italic: true } };
    numberedLayout.slideNumber = { x: 200, align: 'center' };
    master.slideNumber = { x: 300, align: 'right' };
    const duplicate = imported.duplicateSlide(0);
    imported.moveSlide(imported.slides.indexOf(duplicate), 0);
    expect(imported.opcPackage.requirePart(themeUri).bytes).toEqual(themeBefore);
    expect(imported.opcPackage.relationships(imported.slides[1]!.partUri)
      .map((relationship) => ({ ...relationship }))).toEqual(slideRelationshipsBefore);

    const reopened = await PptxDocument.open(await imported.write());
    expect(reopened.slides.map((slide) => slideNumberOwnerState(reopened, slide.partUri).cache))
      .toEqual(['1', '2']);
    expect(reopened.layouts.find(({ partUri }) => partUri === numberedLayout.partUri)?.slideNumber)
      .toMatchObject({ x: 200, align: 'center' });
    expect(reopened.masters[0]?.slideNumber).toMatchObject({ x: 300, align: 'right' });
    expect(slideNumberOwnerState(reopened, reopened.masters[0]!.partUri))
      .toMatchObject({ cache: '‹#›', masterFlag: '1' });
    for (const compatibility of profiles) {
      await reopened.write({ compatibility });
      expect(reopened.diagnostics.filter(({ code }) => code.startsWith('SLIDE_NUMBER_')))
        .toEqual([]);
    }
  });

  it('imports the supported PptxGenJS layout master background matrix', async () => {
    const source = new PptxGenJS();
    const definitions = [
      { title: 'BACKGROUND_SOLID', background: { color: '4472C4' } },
      {
        title: 'BACKGROUND_TRANSPARENT',
        background: { color: 'FF3399', transparency: 50 },
      },
      { title: 'BACKGROUND_IMAGE', background: { data: PNG_DATA_URI } },
      { title: 'BACKGROUND_INHERITED' },
      { title: 'BACKGROUND_DEPRECATED', bkgd: { color: '112233' } },
      { title: 'BACKGROUND_EMPTY', background: {} },
      { title: 'BACKGROUND_INVALID', background: { color: '' } },
    ] as const;
    for (const definition of definitions) {
      source.defineSlideMaster({ ...definition, objects: [] });
      source.addSlide({ masterName: definition.title });
    }
    const imported = await importPptxGenJS(source);
    const layout = (name: string) => imported.layouts.find(({ name: value }) =>
      value === name)!;
    const before = imported.opcPackage.mutations.map((mutation) => ({ ...mutation }));

    expect(layout('BACKGROUND_SOLID').background).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '4472C4' },
    });
    expect(layout('BACKGROUND_TRANSPARENT').background).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF3399' },
      transparency: 50,
    });
    const image = layout('BACKGROUND_IMAGE').background;
    expect(image).toMatchObject({ kind: 'image', contentType: 'image/png' });
    if (image?.kind !== 'image') throw new Error('Expected layout image background');
    expect(inspectRasterImage(image.bytes)).toMatchObject({ width: 2, height: 2 });
    for (const name of [
      'BACKGROUND_INHERITED',
      'BACKGROUND_DEPRECATED',
      'BACKGROUND_EMPTY',
      'BACKGROUND_INVALID',
    ]) {
      expect(layout(name).background, name).toBeUndefined();
    }
    expect(imported.opcPackage.mutations).toEqual(before);
  });

  it('matches public slide master objects, topology, and empty placeholder geometry', async () => {
    const source = new PptxGenJS();
    source.layout = 'LAYOUT_WIDE';
    const placeholderTypes = ['title', 'body', 'pic', 'chart', 'tbl', 'media'] as const;
    const publicObjects: PptxGenJSMasterObject[] = [
      {
        rect: {
          objectName: 'matrix-rect',
          x: 0.5,
          y: 0.5,
          w: 1,
          h: 0.75,
          fill: { color: '112233' },
        },
      },
      {
        line: {
          objectName: 'matrix-line',
          x: 1.75,
          y: 0.5,
          w: 1,
          h: 0.75,
          line: { color: '445566', width: 2 },
        },
      },
      {
        text: {
          text: 'Master text',
          options: {
            objectName: 'matrix-text',
            x: 3,
            y: 0.5,
            w: 2,
            h: 0.75,
            bold: true,
            color: '336699',
          },
        },
      },
      {
        image: {
          objectName: 'matrix-png',
          data: PNG_DATA_URI,
          x: 5.25,
          y: 0.5,
          w: 1,
          h: 0.75,
        },
      },
      {
        image: {
          objectName: 'matrix-svg',
          data: SVG_DATA_URI,
          x: 6.5,
          y: 0.5,
          w: 1,
          h: 0.75,
        },
      },
      ...placeholderTypes.map((type, index): PptxGenJSMasterObject => ({
        placeholder: {
          text: `${type} prompt`,
          options: {
            name: `${type}_matrix`,
            objectName: `${type}-placeholder`,
            type,
            x: 0.5 + index * 2,
            y: 2,
            w: 1.5,
            h: 0.75,
          },
        },
      })),
    ];
    source.defineSlideMaster({ title: 'PUBLIC-MATRIX', objects: publicObjects });
    source.addSlide({ masterName: 'PUBLIC-MATRIX' });
    const imported = await openPptxGenJSPublicOutput(source);

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeObjects: SlideMasterObject[] = [
      {
        kind: 'rect',
        options: {
          name: 'matrix-rect',
          x: inches(0.5),
          y: inches(0.5),
          width: inches(1),
          height: inches(0.75),
          fill: { kind: 'solid', color: { kind: 'srgb', value: '112233' } },
        },
      },
      {
        kind: 'line',
        options: {
          name: 'matrix-line',
          x: inches(1.75),
          y: inches(0.5),
          width: inches(1),
          height: inches(0.75),
          line: { kind: 'line', color: { kind: 'srgb', value: '445566' }, width: 2 },
        },
      },
      {
        kind: 'text',
        text: [{
          runs: [{
            text: 'Master text',
            style: { bold: true, color: { kind: 'srgb', value: '336699' } },
          }],
        }],
        options: {
          name: 'matrix-text',
          x: inches(3),
          y: inches(0.5),
          width: inches(2),
          height: inches(0.75),
        },
      },
      {
        kind: 'image',
        source: PNG_DATA_URI,
        options: {
          name: 'matrix-png',
          x: inches(5.25),
          y: inches(0.5),
          width: inches(1),
          height: inches(0.75),
        },
      },
      {
        kind: 'image',
        source: SVG_DATA_URI,
        options: {
          name: 'matrix-svg',
          fallback: PNG_DATA_URI,
          x: inches(6.5),
          y: inches(0.5),
          width: inches(1),
          height: inches(0.75),
        },
      },
      ...placeholderTypes.map((type, index): SlideMasterObject => ({
        kind: 'placeholder',
        text: `${type} prompt`,
        options: {
          name: `${type}_matrix`,
          type,
          index: 105 + index,
          x: inches(0.5 + index * 2),
          y: inches(2),
          width: inches(1.5),
          height: inches(0.75),
        },
      })),
    ];
    const nativeLayout = await native.defineSlideMaster({
      title: 'NATIVE-MATRIX',
      objects: nativeObjects,
    });
    native.addSlide({ masterName: 'NATIVE-MATRIX' });

    const importedLayout = imported.layouts.find(({ name }) => name === 'PUBLIC-MATRIX')!;
    const importedMaster = imported.masters.find(
      ({ partUri }) => partUri === importedLayout.masterPartUri,
    )!;
    const nativeMaster = native.masters.find(
      ({ partUri }) => partUri === nativeLayout.masterPartUri,
    )!;
    const topology = (
      document: PptxDocument,
      layoutPartUri: string,
      masterPartUri: string,
      slidePartUri: string,
    ) => ({
      layoutBacklinks: document.opcPackage.relationships(layoutPartUri).filter(
        ({ type, resolvedTarget }) =>
          type.endsWith('/slideMaster') && resolvedTarget === masterPartUri,
      ).length,
      masterLinks: document.opcPackage.relationships(masterPartUri).filter(
        ({ type, resolvedTarget }) =>
          type.endsWith('/slideLayout') && resolvedTarget === layoutPartUri,
      ).length,
      slideLinks: document.opcPackage.relationships(slidePartUri).filter(
        ({ type, resolvedTarget }) =>
          type.endsWith('/slideLayout') && resolvedTarget === layoutPartUri,
      ).length,
    });
    expect(topology(
      imported,
      importedLayout.partUri,
      importedMaster.partUri,
      imported.slides[0]!.partUri,
    )).toEqual({ layoutBacklinks: 1, masterLinks: 1, slideLinks: 1 });
    expect(topology(
      native,
      nativeLayout.partUri,
      nativeMaster.partUri,
      native.slides[0]!.partUri,
    )).toEqual({ layoutBacklinks: 1, masterLinks: 1, slideLinks: 1 });

    expect(importedLayout.shapes.map(({ kind }) => kind))
      .toEqual(nativeLayout.shapes.map(({ kind }) => kind));
    expect(importedLayout.shapes.map(({ transform }) => transform))
      .toEqual(nativeLayout.shapes.map(({ transform }) => transform));
    expect(importedLayout.shapes.slice(0, 5).map(({ name }) => name))
      .toEqual(nativeLayout.shapes.slice(0, 5).map(({ name }) => name));
    const importedRect = importedLayout.shapes[0] as ShapeModel;
    const nativeRect = nativeLayout.shapes[0] as ShapeModel;
    const importedLine = importedLayout.shapes[1] as ShapeModel;
    const nativeLine = nativeLayout.shapes[1] as ShapeModel;
    const importedText = importedLayout.shapes[2] as ShapeModel;
    const nativeText = nativeLayout.shapes[2] as ShapeModel;
    expect(importedRect.fill).toEqual(nativeRect.fill);
    expect(importedLine.line).toEqual(nativeLine.line);
    const comparableTextRun = (shape: ShapeModel) => {
      const run = shape.richText[0]?.runs[0];
      return {
        text: run?.text,
        bold: run?.style?.bold,
        color: run?.style?.color,
        lang: run?.style?.lang,
      };
    };
    expect(comparableTextRun(importedText)).toEqual(comparableTextRun(nativeText));

    const importedImages = importedLayout.shapes.filter(
      (shape): shape is ImageModel => shape instanceof ImageModel,
    );
    const nativeImages = nativeLayout.shapes.filter(
      (shape): shape is ImageModel => shape instanceof ImageModel,
    );
    expect(importedImages.map(({ isSvg }) => isSvg)).toEqual([false, true]);
    expect(nativeImages.map(({ isSvg }) => isSvg)).toEqual([false, true]);
    const imageHash = (document: PptxDocument, image: ImageModel) => createHash('sha256')
      .update(document.opcPackage.requirePart(
        image.isSvg ? image.svgPartUri! : image.sourcePartUri!,
      ).bytes)
      .digest('hex');
    expect(importedImages.map((image) => imageHash(imported, image)))
      .toEqual(nativeImages.map((image) => imageHash(native, image)));

    expect(importedLayout.placeholders.map(({ placeholder }) => placeholder)).toEqual([
      { type: 'title', index: 105 },
      { type: 'body', index: 106 },
      { type: 'body', index: 107 },
      { type: 'chart', index: 108 },
      { type: 'body', index: 109 },
      { type: 'media', index: 110 },
    ]);
    expect(nativeLayout.placeholders.map(({ placeholder }) => placeholder)).toEqual(
      placeholderTypes.map((type, index) => ({ type, index: 105 + index })),
    );
    expect(imported.slides[0]?.placeholders.map(({ placeholder }) => placeholder))
      .toEqual(importedLayout.placeholders.map(({ placeholder }) => placeholder));
    expect(native.slides[0]?.placeholders.map(({ placeholder }) => placeholder))
      .toEqual(nativeLayout.placeholders.map(({ placeholder }) => placeholder));

    for (const compatibility of [
      'powerpoint-2010',
      'powerpoint-current',
      'keynote-current',
      'libreoffice-current',
      'google-slides-import',
    ] as const) {
      await imported.write({ compatibility });
      await native.write({ compatibility });
      expect(imported.diagnostics.filter(({ code, severity }) =>
        severity !== 'info' && /^(LAYOUT_|PLACEHOLDER_)/.test(code))).toEqual([]);
      expect(native.diagnostics.filter(({ code, severity }) =>
        severity !== 'info' && /^(LAYOUT_|PLACEHOLDER_)/.test(code))).toEqual([]);
    }
  });

  it('compares PptxGenJS and native chart creation return semantics', async () => {
    const generated = new PptxGenJS();
    const generatedSlide = generated.addSlide();
    const generatedResult = generatedSlide.addChart(generated.ChartType.bar!, [{
      name: 'Revenue', labels: ['Q1', 'Q2'], values: [10, 20],
    }]);
    expect(generatedResult).toBe(generatedSlide);

    const native = PptxDocument.create();
    native.addSlide();
    const nativeResult = native.addChart(0, 'bar', [{
      name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20],
    }]);
    expect(nativeResult).toBeInstanceOf(Promise);
    const chart = await nativeResult;
    expect(chart).toBeInstanceOf(ChartModel);
    expect(native.slides[0]!.shapes).toContain(chart);
  });

  it('matches all public slide master chart types and combo workbooks to native output', async () => {
    const source = new PptxGenJS();
    source.layout = 'LAYOUT_WIDE';
    const chartTypes = [
      ['area', source.ChartType.area!],
      ['bar', source.ChartType.bar!],
      ['bar3D', source.ChartType.bar3d!],
      ['bubble', source.ChartType.bubble!],
      ['doughnut', source.ChartType.doughnut!],
      ['line', source.ChartType.line!],
      ['pie', source.ChartType.pie!],
      ['radar', source.ChartType.radar!],
      ['scatter', source.ChartType.scatter!],
    ] as const;
    const publicData = (type: typeof chartTypes[number][0]): readonly PptxGenJSChartData[] => {
      if (type === 'scatter') {
        return [
          { name: 'X', values: [1, 2] },
          { name: 'scatter-series', values: [10, 20] },
        ];
      }
      if (type === 'bubble') {
        return [
          { name: 'X', values: [1, 2] },
          { name: 'bubble-series', values: [10, 20], sizes: [5, 6] },
        ];
      }
      return [{ name: `${type}-series`, labels: ['Q1', 'Q2'], values: [10, 20] }];
    };
    const publicObjects: PptxGenJSMasterObject[] = chartTypes.map(
      ([type, publicType], index) => ({
        chart: {
          type: publicType,
          data: publicData(type),
          opts: { x: 0.5 + index, y: 0.5, w: 0.8, h: 1.5 },
        },
      }),
    );
    const comboData = [{ name: 'Revenue', labels: ['Q1', 'Q2'], values: [10, 20] }];
    const trendData = [{ name: 'Trend', labels: ['Q1', 'Q2'], values: [11, 21] }];
    publicObjects.push({
      chart: {
        type: [
          { type: source.ChartType.bar!, data: comboData, options: {} },
          {
            type: source.ChartType.line!,
            data: trendData,
            options: { secondaryCatAxis: true, secondaryValAxis: true },
          },
        ],
        data: {
          x: 9.5,
          y: 0.5,
          w: 2,
          h: 1.5,
          catAxes: [{}, {}],
          valAxes: [{}, {}],
        },
        opts: {},
      },
    });
    source.defineSlideMaster({ title: 'PUBLIC-CHARTS', objects: publicObjects });
    source.addSlide({ masterName: 'PUBLIC-CHARTS' });
    const imported = await openPptxGenJSPublicOutput(source);
    const importedLayout = imported.layouts.find(({ name }) => name === 'PUBLIC-CHARTS')!;
    const importedCharts = importedLayout.shapes.filter(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    );

    const nativeObject = (
      type: typeof chartTypes[number][0],
      index: number,
    ): SlideMasterObject => {
      const options = {
        x: inches(0.5 + index),
        y: inches(0.5),
        width: inches(0.8),
        height: inches(1.5),
      };
      if (type === 'scatter') {
        return {
          kind: 'chart',
          groups: [{
            type,
            series: [{ name: 'scatter-series', xValues: [1, 2], values: [10, 20] }],
          }],
          options,
        };
      }
      if (type === 'bubble') {
        return {
          kind: 'chart',
          groups: [{
            type,
            series: [{
              name: 'bubble-series',
              xValues: [1, 2],
              values: [10, 20],
              sizes: [5, 6],
            }],
          }],
          options,
        };
      }
      return {
        kind: 'chart',
        groups: [{
          type,
          series: [{ name: `${type}-series`, categories: ['Q1', 'Q2'], values: [10, 20] }],
        }],
        options,
      };
    };
    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeLayout = await native.defineSlideMaster({
      title: 'NATIVE-CHARTS',
      objects: [
        ...chartTypes.map(([type], index) => nativeObject(type, index)),
        {
          kind: 'chart',
          groups: [
            {
              type: 'bar',
              series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
            },
            {
              type: 'line',
              axis: 'secondary',
              series: [{ name: 'Trend', categories: ['Q1', 'Q2'], values: [11, 21] }],
            },
          ],
          options: {
            x: inches(9.5),
            y: inches(0.5),
            width: inches(2),
            height: inches(1.5),
          },
        },
      ],
    });
    native.addSlide({ masterName: 'NATIVE-CHARTS' });
    const nativeCharts = nativeLayout.shapes.filter(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    );
    expect(importedCharts).toHaveLength(10);
    expect(nativeCharts).toHaveLength(10);
    const chartState = (charts: readonly ChartModel[]) => charts.map((chart) => ({
      transform: chart.transform,
      groups: chart.definition?.groups.map(({ type, axis, series }) => ({
        type,
        axis,
        series: series.map(({ name, values, xValues, sizes }) => ({
          name,
          values,
          xValues,
          sizes,
        })),
      })),
    }));
    expect(chartState(importedCharts)).toEqual(chartState(nativeCharts));
    expect(importedCharts.at(-1)?.definition?.groups.map(({ type, axis }) => [type, axis]))
      .toEqual([['bar', 'primary'], ['line', 'secondary']]);

    for (const [document, layout, charts] of [
      [imported, importedLayout, importedCharts],
      [native, nativeLayout, nativeCharts],
    ] as const) {
      for (const chart of charts) {
        expect(chart.workbookPartUri).toBeDefined();
        expect(await chartWorkbookMatches(
          document.opcPackage.requirePart(chart.workbookPartUri!).bytes,
          chart.definition!,
          chart.xml,
        )).toBe(true);
        expect(document.opcPackage.relationships(layout.partUri).filter(
          ({ type, resolvedTarget }) =>
            type.endsWith('/chart') && resolvedTarget === chart.chartPartUri,
        )).toHaveLength(1);
        expect(document.opcPackage.relationships(chart.chartPartUri!).filter(
          ({ type, resolvedTarget }) =>
            type.endsWith('/package') && resolvedTarget === chart.workbookPartUri,
        )).toHaveLength(1);
      }
    }
  });

  it('compares public placeholder population payloads with strict native owners', async () => {
    const source = new PptxGenJS();
    source.layout = 'LAYOUT_WIDE';
    const publicPlaceholders = [
      {
        name: 'title_box', objectName: 'title-placeholder',
        type: 'title', x: 0.5, y: 0.5, w: 2, h: 0.75,
      },
      {
        name: 'rich_box', objectName: 'rich-placeholder',
        type: 'body', x: 0.5, y: 1.5, w: 2, h: 0.75,
      },
      {
        name: 'pic_box', objectName: 'picture-placeholder',
        type: 'pic', x: 3, y: 0.5, w: 2, h: 1.75,
      },
      {
        name: 'chart_box', objectName: 'chart-placeholder',
        type: 'chart', x: 5.5, y: 0.5, w: 2.5, h: 1.75,
      },
      {
        name: 'table_box', objectName: 'table-placeholder',
        type: 'tbl', x: 8.5, y: 0.5, w: 2.5, h: 1.75,
      },
      {
        name: 'media_box', objectName: 'media-placeholder',
        type: 'media', x: 11.5, y: 0.5, w: 1.25, h: 1.75,
      },
    ] as const;
    source.defineSlideMaster({
      title: 'PUBLIC-POPULATED',
      objects: publicPlaceholders.map((options): PptxGenJSMasterObject => ({
        placeholder: { text: `${options.type} prompt`, options },
      })),
    });
    const publicSlide = source.addSlide({ masterName: 'PUBLIC-POPULATED' });
    publicSlide.addText('Filled title', { placeholder: 'title_box' });
    publicSlide.addText([
      { text: 'Rich', options: { bold: true } },
      { text: ' body', options: { italic: true } },
    ], { placeholder: 'rich_box' });
    publicSlide.addImage({ data: PNG_DATA_URI, placeholder: 'pic_box' });
    publicSlide.addChart(source.ChartType.bar!, [{
      name: 'Revenue', labels: ['Q1', 'Q2'], values: [10, 20],
    }], { placeholder: 'chart_box' });
    publicSlide.addTable([
      [{ text: 'Quarter' }, { text: 'Revenue' }],
      [{ text: 'Q1' }, { text: '10' }],
    ], { placeholder: 'table_box' });
    const publicMediaOptions: PptxGenJSMediaOptions & { readonly placeholder: string } = {
      type: 'audio',
      data: 'data:audio/mpeg;base64,AQIDBA==',
      placeholder: 'media_box',
    };
    publicSlide.addMedia(publicMediaOptions);
    const imported = await openPptxGenJSPublicOutput(source);
    const importedLayout = imported.layouts.find(({ name }) => name === 'PUBLIC-POPULATED')!;
    const importedSlide = imported.slides[0]!;

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeLayout = await native.defineSlideMaster({
      title: 'NATIVE-POPULATED',
      objects: publicPlaceholders.map((placeholder, index): SlideMasterObject => ({
        kind: 'placeholder',
        text: `${placeholder.type} prompt`,
        options: {
          name: placeholder.name,
          type: placeholder.type,
          index: 100 + index,
          x: inches(placeholder.x),
          y: inches(placeholder.y),
          width: inches(placeholder.w),
          height: inches(placeholder.h),
        },
      })),
    });
    const nativeSlide = native.addSlide({ masterName: 'NATIVE-POPULATED' });
    nativeSlide.addText('Filled title', { placeholder: 'title_box' });
    nativeSlide.addRichText([
      {
        runs: [
          { text: 'Rich', style: { bold: true } },
          { text: ' body', style: { italic: true } },
        ],
      },
    ], { placeholder: 'rich_box' });
    await native.addImage(0, PNG_DATA_URI, { placeholder: 'pic_box' });
    await native.addChart(0, 'bar', [{
      name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20],
    }], { placeholder: 'chart_box' });
    nativeSlide.addTable([
      ['Quarter', 'Revenue'],
      ['Q1', '10'],
    ], { placeholder: 'table_box' });
    await native.addAudio(0, 'data:audio/mpeg;base64,AQIDBA==', {
      placeholder: 'media_box',
    });

    expect(importedLayout.placeholders.map(({ placeholder }) => placeholder)).toEqual([
      { type: 'title', index: 100 },
      { type: 'body', index: 101 },
      { type: 'body', index: 102 },
      { type: 'chart', index: 103 },
      { type: 'body', index: 104 },
      { type: 'media', index: 105 },
    ]);
    expect(nativeLayout.placeholders.map(({ placeholder }) => placeholder)).toEqual([
      { type: 'title', index: 100 },
      { type: 'body', index: 101 },
      { type: 'pic', index: 102 },
      { type: 'chart', index: 103 },
      { type: 'tbl', index: 104 },
      { type: 'media', index: 105 },
    ]);
    expect(importedSlide.shapes.map(({ kind }) => kind)).toEqual([
      'text', 'text', 'image', 'chart', 'table', 'audio', 'text',
    ]);
    expect(nativeSlide.shapes.map(({ kind }) => kind)).toEqual([
      'text', 'text', 'image', 'chart', 'table', 'audio',
    ]);
    expect((importedSlide.shapes[0] as ShapeModel).text)
      .toBe((nativeSlide.shapes[0] as ShapeModel).text);
    expect((importedSlide.shapes[1] as ShapeModel).richText[0]?.runs.map(
      ({ text, style }) => ({ text, bold: style?.bold, italic: style?.italic }),
    )).toEqual((nativeSlide.shapes[1] as ShapeModel).richText[0]?.runs.map(
      ({ text, style }) => ({ text, bold: style?.bold, italic: style?.italic }),
    ));

    const importedImage = importedSlide.shapes.find(
      (shape): shape is ImageModel => shape instanceof ImageModel,
    )!;
    const nativeImage = nativeSlide.shapes.find(
      (shape): shape is ImageModel => shape instanceof ImageModel,
    )!;
    expect(importedImage.transform).toMatchObject({
      x: importedLayout.placeholders[2]?.transform.x,
      y: importedLayout.placeholders[2]?.transform.y,
      width: inches(1),
      height: inches(1),
    });
    expect(importedImage.transform).not.toEqual(importedLayout.placeholders[2]?.transform);
    expect(nativeImage.transform).toEqual(nativeLayout.placeholders[2]?.transform);
    expect(createHash('sha256').update(
      imported.opcPackage.requirePart(importedImage.sourcePartUri!).bytes,
    ).digest('hex')).toBe(createHash('sha256').update(
      native.opcPackage.requirePart(nativeImage.sourcePartUri!).bytes,
    ).digest('hex'));

    const importedChart = importedSlide.shapes.find(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    )!;
    const nativeChart = nativeSlide.shapes.find(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    )!;
    expect(importedChart.transform).toEqual(importedLayout.placeholders[3]?.transform);
    expect(nativeChart.transform).toEqual(nativeLayout.placeholders[3]?.transform);
    expect(importedChart.definition?.groups.map(({ type, series }) => ({
      type,
      series: series.map(({ name, values }) => ({ name, values })),
    }))).toEqual(nativeChart.definition?.groups.map(({ type, series }) => ({
      type,
      series: series.map(({ name, values }) => ({ name, values })),
    })));
    for (const [document, chart] of [
      [imported, importedChart],
      [native, nativeChart],
    ] as const) {
      expect(await chartWorkbookMatches(
        document.opcPackage.requirePart(chart.workbookPartUri!).bytes,
        chart.definition!,
        chart.xml,
      )).toBe(true);
    }

    const importedTable = importedSlide.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    )!;
    const nativeTable = nativeSlide.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    )!;
    const expectedTableRows = [
      ['Quarter', 'Revenue'],
      ['Q1', '10'],
    ];
    expect(importedTable.id).toBe(importedSlide.shapes[0]?.id);
    expect(importedTable.rows).toEqual([]);
    for (const text of expectedTableRows.flat()) {
      expect(slideXml(imported, 0)).toContain(`<a:t>${text}</a:t>`);
    }
    expect(nativeTable.rows.map(({ cells }) => cells.map(({ text }) => text)))
      .toEqual(expectedTableRows);
    const importedAudio = embeddedMediaStates(imported, 0)[0]!;
    const nativeAudio = embeddedMediaStates(native, 0)[0]!;
    expect(importedAudio.mediaBytes).toEqual(nativeAudio.mediaBytes);
    const nativeMediaGeometry = {
      x: nativeLayout.placeholders[5]!.transform.x,
      y: nativeLayout.placeholders[5]!.transform.y,
      width: nativeLayout.placeholders[5]!.transform.width,
      height: nativeLayout.placeholders[5]!.transform.height,
    };
    expect(importedAudio.transform).not.toEqual(nativeMediaGeometry);
    expect(nativeAudio.transform).toEqual(nativeMediaGeometry);
    expect(nativeSlide.placeholders.map(({ placeholder }) => placeholder))
      .toEqual(nativeLayout.placeholders.map(({ placeholder }) => placeholder));

    for (const compatibility of [
      'powerpoint-2010',
      'powerpoint-current',
      'keynote-current',
      'libreoffice-current',
      'google-slides-import',
    ] as const) {
      await imported.write({ mode: 'permissive', compatibility });
      expect(imported.diagnostics.filter(({ code }) => /^(LAYOUT_|PLACEHOLDER_)/.test(code)))
        .toEqual([
          expect.objectContaining({
            severity: 'error',
            code: 'PLACEHOLDER_DOMAIN_MISMATCH',
            partUri: importedSlide.partUri,
            objectId: 'body:102',
            compatibility,
          }),
          expect.objectContaining({
            severity: 'warning',
            code: 'PLACEHOLDER_OWNER_MISSING',
            partUri: importedSlide.partUri,
            objectId: 'body:104',
            compatibility,
          }),
        ]);
      await native.write({ compatibility });
      expect(native.diagnostics.filter(({ code, severity }) =>
        severity !== 'info' && /^(LAYOUT_|PLACEHOLDER_)/.test(code))).toEqual([]);
    }
  });

  it('locks PlaceholderProps and text/image placeholder population against PptxGenJS 4.0.1', async () => {
    const source = new PptxGenJS();
    source.layout = 'LAYOUT_WIDE';
    source.defineSlideMaster({
      title: 'PUBLIC-PLACEHOLDER-CORE',
      objects: [
        {
          placeholder: {
            text: 'Numeric title',
            options: {
              name: 'runtime_title',
              objectName: 'Selection title',
              type: 'title',
              x: 1,
              y: 2,
              w: 3,
              h: 1,
            },
          },
        },
        {
          placeholder: {
            text: 'Percentage body',
            options: {
              name: 'percentage_body',
              objectName: 'Selection percentage body',
              type: 'body',
              x: '10%',
              y: '20%',
              w: '30%',
              h: '40%',
            },
          },
        },
        {
          placeholder: {
            text: 'Picture',
            options: {
              name: 'picture_box',
              objectName: 'Selection picture',
              type: 'pic',
              x: 5,
              y: 1,
              w: 4,
              h: 3,
            },
          },
        },
        {
          placeholder: {
            text: 'Official fields only',
            options: { name: 'official_only', type: 'body' },
          },
        },
      ],
    });
    const publicSlide = source.addSlide({ masterName: 'PUBLIC-PLACEHOLDER-CORE' });
    publicSlide.addText('Filled numeric title', {
      placeholder: 'runtime_title', x: 9, y: 6, w: 2, h: 0.5,
    });
    publicSlide.addText('Filled official fields', {
      placeholder: 'official_only', x: 8, y: 5, w: 2, h: 0.5,
    });
    publicSlide.addText('Unknown selector text', {
      placeholder: 'missing_text', x: 7, y: 6, w: 2, h: 0.5,
    });
    publicSlide.addImage({
      data: PNG_DATA_URI,
      placeholder: 'picture_box',
      x: 9,
      y: 6,
      w: 2,
      h: 0.5,
    });
    publicSlide.addImage({
      data: PNG_DATA_URI,
      placeholder: 'missing_image',
      x: 10,
      y: 6,
      w: 2,
      h: 0.5,
    });

    const imported = await openPptxGenJSPublicOutput(source);
    const importedLayout = imported.layouts.find(
      ({ name }) => name === 'PUBLIC-PLACEHOLDER-CORE',
    )!;
    expect(importedLayout.placeholders.map(({ name, placeholder }) => ({ name, placeholder })))
      .toEqual([
        { name: 'Selection title', placeholder: { type: 'title', index: 100 } },
        { name: 'Selection percentage body', placeholder: { type: 'body', index: 101 } },
        { name: 'Selection picture', placeholder: { type: 'body', index: 102 } },
        { name: 'Text 0', placeholder: { type: 'body', index: 103 } },
      ]);
    expect(importedLayout.placeholders[1]?.transform).toEqual({
      x: 1_219_200,
      y: 1_371_600,
      width: 3_657_600,
      height: 2_743_200,
      rotation: 0,
      flipHorizontal: false,
      flipVertical: false,
    });
    expect(importedLayout.placeholders[3]?.transform).toMatchObject({
      x: 0,
      y: 0,
      width: 9_144_000,
      height: 0,
    });
    const importedTitle = imported.slides[0]!.shapes.find(
      (shape) => shape instanceof ShapeModel && shape.text === 'Filled numeric title',
    ) as ShapeModel;
    const importedOfficial = imported.slides[0]!.shapes.find(
      (shape) => shape instanceof ShapeModel && shape.text === 'Filled official fields',
    ) as ShapeModel;
    const importedUnknownText = imported.slides[0]!.shapes.find(
      (shape) => shape instanceof ShapeModel && shape.text === 'Unknown selector text',
    ) as ShapeModel;
    const importedImages = imported.slides[0]!.shapes.filter(
      (shape): shape is ImageModel => shape instanceof ImageModel,
    );
    const importedPicture = importedImages.find(({ placeholder }) => placeholder !== undefined)!;
    const importedUnknownImage = importedImages.find(
      ({ placeholder }) => placeholder === undefined,
    )!;
    expect({
      name: importedTitle.name,
      placeholder: importedTitle.placeholder,
      transform: importedTitle.transform,
    }).toEqual({
      name: 'Selection title',
      placeholder: { type: 'title', index: 100 },
      transform: importedLayout.placeholders[0]?.transform,
    });
    expect({
      name: importedOfficial.name,
      placeholder: importedOfficial.placeholder,
      transform: importedOfficial.transform,
    }).toEqual({
      name: 'Text 0',
      placeholder: { type: 'body', index: 103 },
      transform: {
        x: inches(8),
        y: inches(5),
        width: inches(2),
        height: inches(0.5),
        rotation: 0,
        flipHorizontal: false,
        flipVertical: false,
      },
    });
    expect(importedUnknownText.placeholder).toBeUndefined();
    expect(importedUnknownText.transform).toMatchObject({
      x: inches(7), y: inches(6), width: inches(2), height: inches(0.5),
    });
    expect({
      name: importedPicture.name,
      placeholder: importedPicture.placeholder,
      transform: importedPicture.transform,
    }).toMatchObject({
      name: 'Image 0',
      placeholder: { type: 'body', index: 102 },
      transform: {
        x: inches(5), y: inches(1), width: inches(2), height: inches(0.5),
      },
    });
    expect(importedUnknownImage.placeholder).toBeUndefined();
    expect(importedUnknownImage.transform).toMatchObject({
      x: inches(10), y: inches(6), width: inches(2), height: inches(0.5),
    });
    await imported.write({ mode: 'permissive', compatibility: 'powerpoint-current' });
    expect(imported.diagnostics).toEqual(expect.arrayContaining([
      expect.objectContaining({
        code: 'PLACEHOLDER_DOMAIN_MISMATCH',
        objectId: 'body:102',
      }),
    ]));

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeLayout = await native.defineSlideMaster({
      title: 'NATIVE-PLACEHOLDER-CORE',
      objects: [
        {
          kind: 'placeholder', text: 'Numeric title',
          options: {
            name: 'runtime_title', type: 'title', index: 100,
            x: inches(1), y: inches(2), width: inches(3), height: inches(1),
          },
        },
        {
          kind: 'placeholder', text: 'Percentage body',
          options: {
            name: 'percentage_body', type: 'body', index: 101,
            x: '10%', y: '20%', width: '30%', height: '40%',
          },
        },
        {
          kind: 'placeholder', text: 'Picture',
          options: {
            name: 'picture_box', type: 'pic', index: 102,
            x: inches(5), y: inches(1), width: inches(4), height: inches(3),
          },
        },
        {
          kind: 'placeholder', text: 'Official fields only',
          options: { name: 'official_only', type: 'body', index: 103 },
        },
      ],
    });
    const nativeSlide = native.addSlide({ masterName: 'NATIVE-PLACEHOLDER-CORE' });
    const nativeTitle = nativeSlide.addText('Filled numeric title', {
      placeholder: 'runtime_title',
      x: inches(9), y: inches(6), width: inches(2), height: inches(0.5),
    });
    const nativeOfficial = nativeSlide.addText('Filled official fields', {
      placeholder: { type: 'body', index: 103 },
      x: inches(8), y: inches(5), width: inches(2), height: inches(0.5),
    });
    const nativePicture = await native.addImage(0, PNG_DATA_URI, {
      placeholder: 'picture_box',
      x: inches(9), y: inches(6), width: inches(2), height: inches(0.5),
    });
    expect(nativeLayout.placeholders.map(({ name, placeholder }) => ({ name, placeholder })))
      .toEqual([
        { name: 'runtime_title', placeholder: { type: 'title', index: 100 } },
        { name: 'percentage_body', placeholder: { type: 'body', index: 101 } },
        { name: 'picture_box', placeholder: { type: 'pic', index: 102 } },
        { name: 'official_only', placeholder: { type: 'body', index: 103 } },
      ]);
    expect(nativeLayout.placeholders[1]?.transform).toEqual(
      importedLayout.placeholders[1]?.transform,
    );
    expect(nativeTitle.transform).toEqual(nativeLayout.placeholders[0]?.transform);
    expect(nativeOfficial.transform).toEqual(nativeLayout.placeholders[3]?.transform);
    expect({
      name: nativePicture.name,
      placeholder: nativePicture.placeholder,
      transform: nativePicture.transform,
    }).toEqual({
      name: 'picture_box',
      placeholder: { type: 'pic', index: 102 },
      transform: nativeLayout.placeholders[2]?.transform,
    });

    const beforeUnknownText = packageState(native);
    expect(() => nativeSlide.addText('Unknown selector text', {
      placeholder: 'missing_text',
    })).toThrow(/placeholder/i);
    expect(packageState(native)).toEqual(beforeUnknownText);
    const beforeUnknownImage = packageState(native);
    await expect(native.addImage(0, PNG_DATA_URI, { placeholder: 'missing_image' }))
      .rejects.toThrow(/placeholder/i);
    expect(packageState(native)).toEqual(beforeUnknownImage);
  });

  it('locks public slide master fallbacks while native definitions reject atomically', async () => {
    const unknown = new PptxGenJS();
    unknown.defineSlideMaster({ title: 'KNOWN', objects: [] });
    unknown.addSlide({ masterName: 'UNKNOWN' });
    const importedUnknown = await openPptxGenJSPublicOutput(unknown);
    const knownLayout = importedUnknown.layouts.find(({ name }) => name === 'KNOWN')!;
    expect(importedUnknown.slides[0]?.relationships.find(
      ({ type }) => type.endsWith('/slideLayout'),
    )?.resolvedTarget).not.toBe(knownLayout.partUri);

    const duplicate = new PptxGenJS();
    duplicate.defineSlideMaster({
      title: 'DUPLICATE',
      objects: [{ text: { text: 'First', options: {} } }],
    });
    duplicate.defineSlideMaster({
      title: 'DUPLICATE',
      objects: [{ text: { text: 'Second', options: {} } }],
    });
    duplicate.addSlide({ masterName: 'DUPLICATE' });
    const importedDuplicate = await openPptxGenJSPublicOutput(duplicate);
    const duplicateLayouts = importedDuplicate.layouts.filter(({ name }) => name === 'DUPLICATE');
    expect(duplicateLayouts).toHaveLength(2);
    expect(importedDuplicate.slides[0]?.relationships.find(
      ({ type }) => type.endsWith('/slideLayout'),
    )?.resolvedTarget).toBe(duplicateLayouts[0]?.partUri);
    await importedDuplicate.write({ mode: 'permissive' });
    expect(importedDuplicate.diagnostics).toEqual(expect.arrayContaining([
      expect.objectContaining({
        severity: 'error',
        code: 'LAYOUT_NAME_DUPLICATE',
        partUri: duplicateLayouts[1]?.partUri,
        objectId: 'DUPLICATE',
      }),
    ]));

    const invalidType = new PptxGenJS();
    invalidType.defineSlideMaster({
      title: 'INVALID-TYPE',
      objects: [{
        placeholder: {
          text: 'Invalid',
          options: { name: 'invalid', type: 'footer', x: 1, y: 1, w: 2, h: 1 },
        },
      }],
    } as never);
    invalidType.addSlide({ masterName: 'INVALID-TYPE' });
    const importedInvalidType = await openPptxGenJSPublicOutput(invalidType);
    expect(importedInvalidType.layouts.find(({ name }) => name === 'INVALID-TYPE')
      ?.placeholders[0]?.placeholder).toEqual({ type: 'body', index: 100 });

    const duplicateNames = new PptxGenJS();
    duplicateNames.defineSlideMaster({
      title: 'DUPLICATE-PLACEHOLDER-NAME',
      objects: [
        {
          placeholder: {
            text: 'First',
            options: { name: 'same_name', type: 'title', x: 1, y: 1, w: 2, h: 1 },
          },
        },
        {
          placeholder: {
            text: 'Second',
            options: { name: 'same_name', type: 'body', x: 1, y: 2, w: 2, h: 1 },
          },
        },
      ],
    });
    duplicateNames.addSlide({ masterName: 'DUPLICATE-PLACEHOLDER-NAME' });
    const importedDuplicateNames = await openPptxGenJSPublicOutput(duplicateNames);
    expect(importedDuplicateNames.slides[0]?.placeholders.map(({ placeholder }) => placeholder))
      .toEqual([{ type: 'title', index: 100 }]);
    await importedDuplicateNames.write({ mode: 'permissive' });
    expect(importedDuplicateNames.diagnostics).toEqual(expect.arrayContaining([
      expect.objectContaining({
        severity: 'error',
        code: 'PLACEHOLDER_IDENTITY_AMBIGUOUS',
        partUri: importedDuplicateNames.layouts.find(
          ({ name }) => name === 'DUPLICATE-PLACEHOLDER-NAME',
        )?.partUri,
        objectId: 'name:Text 0',
      }),
    ]));

    const zeroGeometry = new PptxGenJS();
    zeroGeometry.defineSlideMaster({
      title: 'ZERO-GEOMETRY',
      objects: [{
        placeholder: {
          text: 'Zero',
          options: { name: 'zero_box', type: 'title', x: 0, y: 0, w: 0, h: 0 },
        },
      }],
    });
    zeroGeometry.addSlide({ masterName: 'ZERO-GEOMETRY' });
    const importedZero = await openPptxGenJSPublicOutput(zeroGeometry);
    expect(importedZero.layouts.find(({ name }) => name === 'ZERO-GEOMETRY')
      ?.placeholders[0]?.transform).toMatchObject({ x: 0, y: 0, width: 0, height: 0 });
    expect(importedZero.slides[0]?.placeholders[0]?.transform)
      .toMatchObject({ x: 0, y: 0, width: 0, height: 0 });

    const delayed = new PptxGenJS();
    delayed.defineSlideMaster({
      title: 'DELAYED',
      objects: [{
        placeholder: {
          text: 'Delayed',
          options: { name: 'delayed_box', type: 'title', x: 1, y: 1, w: 2, h: 1 },
        },
      }],
    });
    const delayedSlide = delayed.addSlide({ masterName: 'DELAYED' });
    const firstWrite = await openPptxGenJSPublicOutput(delayed);
    expect(firstWrite.slides[0]?.placeholders.map((shape) => ({
      placeholder: shape.placeholder,
      text: (shape as ShapeModel).text,
    }))).toEqual([{ placeholder: { type: 'title', index: 100 }, text: '' }]);
    delayedSlide.addText('Filled after first write', { placeholder: 'delayed_box' });
    const secondWrite = await openPptxGenJSPublicOutput(delayed);
    expect(secondWrite.slides[0]?.placeholders.map((shape) => ({
      placeholder: shape.placeholder,
      text: (shape as ShapeModel).text,
    }))).toEqual([
      { placeholder: { type: 'title', index: 100 }, text: '' },
      { placeholder: { type: 'title', index: 100 }, text: 'Filled after first write' },
    ]);
    await secondWrite.write({ mode: 'permissive' });
    expect(secondWrite.diagnostics).toEqual(expect.arrayContaining([
      expect.objectContaining({
        severity: 'error',
        code: 'PLACEHOLDER_IDENTITY_AMBIGUOUS',
        partUri: secondWrite.slides[0]?.partUri,
        objectId: 'title:100',
      }),
    ]));

    const nativeUnknown = PptxDocument.create();
    const unknownBefore = packageState(nativeUnknown);
    expect(() => nativeUnknown.addSlide({ masterName: 'UNKNOWN' })).toThrow(/not found/i);
    expect(packageState(nativeUnknown)).toEqual(unknownBefore);

    const nativeDuplicate = PptxDocument.create();
    await nativeDuplicate.defineSlideMaster({ title: 'DUPLICATE' });
    const duplicateBefore = packageState(nativeDuplicate);
    await expect(nativeDuplicate.defineSlideMaster({ title: 'DUPLICATE' }))
      .rejects.toThrow(/already|unique/i);
    expect(packageState(nativeDuplicate)).toEqual(duplicateBefore);

    const rejectedDefinitions = [
      {
        title: 'INVALID-TYPE',
        objects: [{
          kind: 'placeholder',
          options: { name: 'invalid', type: 'footer' },
        }],
      },
      {
        title: 'DUPLICATE-PLACEHOLDER-NAME',
        objects: [
          {
            kind: 'placeholder',
            options: { name: 'same_name', type: 'title', index: 100 },
          },
          {
            kind: 'placeholder',
            options: { name: 'same_name', type: 'body', index: 101 },
          },
        ],
      },
      {
        title: 'ZERO-GEOMETRY',
        objects: [{
          kind: 'placeholder',
          options: { name: 'zero_box', type: 'title', width: 0, height: 0 },
        }],
      },
    ] as const;
    for (const definition of rejectedDefinitions) {
      const document = PptxDocument.create();
      const before = packageState(document);
      await expect(document.defineSlideMaster(definition as never)).rejects.toThrow();
      expect(packageState(document)).toEqual(before);
    }

    const nativeDelayed = PptxDocument.create();
    await nativeDelayed.defineSlideMaster({
      title: 'DELAYED',
      objects: [{
        kind: 'placeholder',
        text: 'Delayed',
        options: { name: 'delayed_box', type: 'title', index: 100 },
      }],
    });
    const nativeDelayedSlide = nativeDelayed.addSlide({ masterName: 'DELAYED' });
    await nativeDelayed.write();
    nativeDelayedSlide.addText('Filled after first write', { placeholder: 'delayed_box' });
    expect(nativeDelayedSlide.placeholders.map((shape) => ({
      placeholder: shape.placeholder,
      text: (shape as ShapeModel).text,
    }))).toEqual([{
      placeholder: { type: 'title', index: 100 },
      text: 'Filled after first write',
    }]);
    await nativeDelayed.write();
    expect(nativeDelayed.diagnostics.filter(({ code }) => /^(LAYOUT_|PLACEHOLDER_)/.test(code)))
      .toEqual([]);
  });

  it('reports a real fixed-id collision created only through PptxGenJS public APIs', async () => {
    const source = new PptxGenJS();
    const slide = source.addSlide();
    for (let index = 0; index < 24; index += 1) {
      slide.addText(String(index), { x: 0, y: 0, w: 1, h: 0.2 });
    }
    slide.slideNumber = {};
    const imported = await importPptxGenJS(source);
    expect(slideNumberOwnerState(imported, imported.slides[0]!.partUri))
      .toMatchObject({ shapeId: 25, shapeIdOccurrences: 2 });
    await expect(imported.write()).rejects.toBeInstanceOf(ValidationError);
    expect(imported.diagnostics).toEqual(expect.arrayContaining([
      expect.objectContaining({
        severity: 'error',
        code: 'SLIDE_NUMBER_SHAPE_ID_COLLISION',
        partUri: imported.slides[0]!.partUri,
      }),
    ]));
    await expect(imported.write({ mode: 'permissive' })).resolves.toBeInstanceOf(Uint8Array);
  });

  it('matches supported public PptxGenJS slide backgrounds and locks none divergences', async () => {
    const generated = new PptxGenJS();
    generated.addSlide();
    generated.addSlide().background = { type: 'none' };
    generated.addSlide().background = { color: 'FF3399' };
    generated.addSlide().background = { color: 'FF3399', transparency: 50 };
    generated.addSlide().background = { data: PNG_DATA_URI };
    generated.addSlide().background = { fill: '00FF00' };
    generated.addSlide().background = { type: 'none', color: 'FF3399' };
    const imported = await openPptxGenJSPublicOutput(generated);

    expect(imported.slides.map(({ background }) => background?.kind)).toEqual([
      undefined,
      undefined,
      'solid',
      'solid',
      'image',
      'solid',
      undefined,
    ]);
    expect(imported.slides[2]!.background).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF3399' },
    });
    expect(imported.slides[3]!.background).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF3399' },
      transparency: 50,
    });
    expect(imported.slides[5]!.background).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '00FF00' },
    });

    const native = PptxDocument.create();
    native.addSlide();
    native.addSlide().background = { kind: 'none' };
    native.addSlide().background = {
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF3399' },
    };
    native.addSlide().background = {
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF3399' },
      transparency: 50,
    };
    native.addSlide();
    await native.setSlideBackgroundImage(4, PNG_DATA_URI);
    native.addSlide().background = {
      kind: 'solid',
      color: { kind: 'srgb', value: '00FF00' },
    };

    expect(slideBackgroundStructuralState(imported, 0))
      .toEqual(slideBackgroundStructuralState(native, 0));
    for (const index of [2, 3, 4, 5]) {
      expect(slideBackgroundStructuralState(imported, index))
        .toEqual(slideBackgroundStructuralState(native, index));
    }
    expect(slideBackgroundStructuralState(imported, 1)).toMatchObject({
      kind: undefined,
      direct: { present: false, noFill: false },
    });
    expect(slideBackgroundStructuralState(native, 1)).toMatchObject({
      kind: 'none',
      direct: { present: true, noFill: true },
    });
    expect(slideBackgroundStructuralState(imported, 6)).toMatchObject({
      kind: undefined,
      direct: {
        present: true,
        noFill: false,
        solidFill: false,
        blipFill: false,
      },
    });
    expect(slideXml(imported, 6)).toContain('<p:bg><p:bgPr></p:bgPr></p:bg>');
  });

  it('edits and validates imported PptxGenJS backgrounds without disturbing neighbors', async () => {
    const generated = new PptxGenJS();
    const solidSlide = generated.addSlide();
    solidSlide.background = { color: 'FF3399', transparency: 50 };
    solidSlide.addText('Keep solid neighbor', { x: 1, y: 1, w: 4, h: 1 });
    const imageSlide = generated.addSlide();
    imageSlide.background = { data: PNG_DATA_URI };
    imageSlide.addText('Keep image neighbor', { x: 1, y: 1, w: 4, h: 1 });

    let imported = await importPptxGenJS(generated);
    const unrelatedRelationships = imported.slides.map((slide) =>
      slide.relationships
        .filter(({ type }) => type !== IMAGE_RELATIONSHIP)
        .map(({ id, type, target, targetMode }) => ({ id, type, target, targetMode })));
    const originalImageTarget = imported.slides[1]!.relationships.find(({ type }) =>
      type === IMAGE_RELATIONSHIP)!.resolvedTarget!;

    await imported.setSlideBackgroundImage(0, PNG_DATA_URI);
    imported = await PptxDocument.open(await imported.write());
    const firstImageTarget = imported.slides[0]!.relationships.find(({ type }) =>
      type === IMAGE_RELATIONSHIP)!.resolvedTarget!;
    expect(imported.slides[0]!.background?.kind).toBe('image');
    expect(slideXml(imported, 0)).toContain('Keep solid neighbor');

    imported.slides[0]!.background = {
      kind: 'linear-gradient',
      angle: 45,
      stops: [
        { offset: 0, color: 'FFFFFF' },
        { offset: 1, color: '000000' },
      ],
    };
    imported = await PptxDocument.open(await imported.write());
    expect(imported.slides[0]!.background?.kind).toBe('linear-gradient');
    expect(imported.opcPackage.hasPart(firstImageTarget)).toBe(false);

    imported.slides[0]!.background = undefined;
    imported = await PptxDocument.open(await imported.write());
    expect(imported.slides[0]!.background).toBeUndefined();
    expect(slideXml(imported, 0)).toContain('Keep solid neighbor');

    imported.slides[1]!.background = {
      kind: 'solid',
      color: { kind: 'srgb', value: '00FF00' },
    };
    imported = await PptxDocument.open(await imported.write());
    expect(imported.slides[1]!.background).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '00FF00' },
    });
    expect(imported.opcPackage.hasPart(originalImageTarget)).toBe(false);
    expect(slideXml(imported, 1)).toContain('Keep image neighbor');
    expect(imported.slides.map((slide) =>
      slide.relationships
        .filter(({ type }) => type !== IMAGE_RELATIONSHIP)
        .map(({ id, type, target, targetMode }) => ({ id, type, target, targetMode }))))
      .toEqual(unrelatedRelationships);

    for (const compatibility of [
      'powerpoint-2010',
      'powerpoint-current',
      'keynote-current',
      'google-slides-import',
      'libreoffice-current',
    ] as const) {
      await imported.write({ mode: 'permissive', compatibility });
      expect(imported.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
    }
  });

  it('normalizes and semantically replaces imported PptxGenJS chart families', async () => {
    const generated = new PptxGenJS();
    const generatedSlide = generated.addSlide();
    generatedSlide.addChart(generated.ChartType.bar!, [{
      name: 'Revenue', labels: ['Q1', 'Q2'], values: [10, 20],
    }], {
      x: 1,
      y: 1,
      w: 6,
      h: 4,
      showLegend: false,
    });
    generatedSlide.addChart(generated.ChartType.area!, [{
      name: 'Pipeline', labels: ['Q1', 'Q2'], values: [8, 16],
    }], { x: 1, y: 1, w: 6, h: 4 });
    generatedSlide.addChart(generated.ChartType.pie!, [{
      name: 'Share', labels: ['A', 'B'], values: [60, 40],
    }], { x: 1, y: 1, w: 6, h: 4 });
    generatedSlide.addChart(generated.ChartType.scatter!, [
      { name: 'X', values: [1, 2] },
      { name: 'Points', values: [3, 4] },
    ], { x: 1, y: 1, w: 6, h: 4 });
    generatedSlide.addChart(generated.ChartType.bubble!, [
      { name: 'X', values: [1, 2] },
      { name: 'Bubbles', values: [3, 4], sizes: [5, 6] },
    ], { x: 1, y: 1, w: 6, h: 4 });
    generatedSlide.addChart([
      {
        type: generated.ChartType.bar!,
        data: [{ name: 'Revenue', labels: ['Q1', 'Q2'], values: [10, 20] }],
        options: {},
      },
      {
        type: generated.ChartType.line!,
        data: [{ name: 'Trend', labels: ['Q1', 'Q2'], values: [11, 21] }],
        options: { secondaryCatAxis: true, secondaryValAxis: true },
      },
    ], {
      x: 1,
      y: 1,
      w: 6,
      h: 4,
      catAxes: [{}, {}],
      valAxes: [{}, {}],
    });

    const imported = await importPptxGenJS(generated);
    const charts = imported.slides[0]!.shapes.filter(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    );
    expect(charts).toHaveLength(6);
    const [bar, area, pie, scatter, bubble, combo] = charts as [
      ChartModel, ChartModel, ChartModel, ChartModel, ChartModel, ChartModel,
    ];
    expect(bar.definition).toMatchObject({
      groups: [{
        type: 'bar',
        axis: 'primary',
        series: [{ name: 'Revenue', categories: [['Q1', 'Q2']], values: [10, 20] }],
      }],
    });
    expect(area.definition).toMatchObject({
      groups: [{ type: 'area', series: [{ name: 'Pipeline', values: [8, 16] }] }],
    });
    expect(pie.definition).toMatchObject({
      groups: [{ type: 'pie', series: [{ name: 'Share', values: [60, 40] }] }],
    });
    expect(scatter.definition).toMatchObject({
      groups: [{
        type: 'scatter',
        axis: 'primary',
        series: [{ name: 'Points', xValues: [1, 2], values: [3, 4] }],
      }],
    });
    expect(bubble.definition).toMatchObject({
      groups: [{
        type: 'bubble',
        axis: 'primary',
        series: [{ name: 'Bubbles', xValues: [1, 2], values: [3, 4], sizes: [5, 6] }],
      }],
    });
    expect(combo.definition?.groups.map(({ type, axis }) => [type, axis])).toEqual([
      ['bar', 'primary'],
      ['line', 'secondary'],
    ]);
    expect(Object.isFrozen(bar.definition?.options)).toBe(true);
    expect(bar.definition?.groups[0]?.options?.series?.[0]?.fill).toBeDefined();
    const importedStyle = /<c:ser>[\s\S]*?(<c:spPr>[\s\S]*?<\/c:spPr>)/
      .exec(bar.xml)?.[1];
    expect(importedStyle).toBeDefined();
    for (const chart of charts) {
      expect(chart.xml).toMatch(/<c:f>[^<]*Sheet1!/);
      expect(await chartWorkbookMatches(
        imported.opcPackage.requirePart(chart.workbookPartUri!).bytes,
        chart.definition!,
        chart.xml,
      )).toBe(true);
      expect(imported.slides[0]!.relationships.filter(
        ({ resolvedTarget }) => resolvedTarget === chart.chartPartUri,
      )).toHaveLength(1);
      expect(imported.opcPackage.relationships(chart.chartPartUri!).filter(
        ({ type, resolvedTarget }) =>
          type.endsWith('/package') && resolvedTarget === chart.workbookPartUri,
      )).toHaveLength(1);
    }
    await expect(imported.write()).resolves.toBeInstanceOf(Uint8Array);
    expect(imported.diagnostics.filter(({ code }) => code.startsWith('CHART_'))).toEqual([]);

    await bar.replaceSeries([{
      name: 'Revenue edited',
      categories: ['Q1', 'Q2', 'Q3'],
      values: [12, 24, 36],
    }]);
    await area.replaceSeries([{
      name: 'Pipeline edited', categories: ['Q1', 'Q2'], values: [9, 18],
    }]);
    await pie.replaceSeries([{
      name: 'Share edited', categories: ['A', 'B'], values: [55, 45],
    }]);
    await scatter.replaceSeries([{
      name: 'Points edited', xValues: [2, 4], values: [6, 8],
    }]);
    await bubble.replaceSeries([{
      name: 'Bubbles edited', xValues: [2, 4], values: [6, 8], sizes: [10, 12],
    }]);
    const comboDefinition = combo.definition!;
    const comboBarGroup = comboDefinition.groups[0];
    const comboLineGroup = comboDefinition.groups[1];
    if (comboBarGroup?.type !== 'bar' || comboLineGroup?.type !== 'line') {
      throw new Error('Expected imported bar-line combination chart');
    }
    await combo.replaceDefinition({ groups: [
      {
        type: 'bar',
        series: [{ name: 'Revenue edited', categories: ['Q1', 'Q2'], values: [12, 22] }],
        ...(comboBarGroup.options === undefined ? {} : { options: comboBarGroup.options }),
      },
      {
        type: 'line',
        axis: 'secondary',
        series: [{ name: 'Trend edited', categories: ['Q1', 'Q2'], values: [13, 23] }],
        ...(comboLineGroup.options === undefined ? {} : { options: comboLineGroup.options }),
      },
    ], options: comboDefinition.options });
    expect(bar.series).toEqual([{
      name: 'Revenue edited',
      categories: ['Q1', 'Q2', 'Q3'],
      values: [12, 24, 36],
    }]);
    expect(bar.xml).toContain(importedStyle!);
    expect(area.series[0]).toMatchObject({ name: 'Pipeline edited', values: [9, 18] });
    expect(pie.series[0]).toMatchObject({ name: 'Share edited', values: [55, 45] });
    expect(scatter.series[0]).toMatchObject({ xValues: [2, 4], values: [6, 8] });
    expect(bubble.series[0]).toMatchObject({ sizes: [10, 12] });
    expect(combo.series.map(({ name, values }) => ({ name, values }))).toEqual([
      { name: 'Revenue edited', values: [12, 22] },
      { name: 'Trend edited', values: [13, 23] },
    ]);

    const reopened = await PptxDocument.open(await imported.write());
    const reopenedCharts = reopened.slides[0]!.shapes.filter(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    );
    expect(reopenedCharts.map(({ definition }) => definition)).toEqual(
      charts.map(({ definition }) => definition),
    );
    expect(reopenedCharts.every(({ workbookPartUri }) => workbookPartUri !== undefined)).toBe(true);
    for (const chart of reopenedCharts) {
      expect(await chartWorkbookMatches(
        reopened.opcPackage.requirePart(chart.workbookPartUri!).bytes,
        chart.definition!,
      )).toBe(true);
    }
  });

  it('locks chart area and plot area fill, border, rounding, and deprecated aliases', async () => {
    const generated = new PptxGenJS();
    const slide = generated.addSlide();
    const data = [{ name: 'Revenue', labels: ['Q1', 'Q2'], values: [10, 20] }];
    const geometry = { x: 1, y: 1, w: 6, h: 4 };

    slide.addChart(generated.ChartType.bar!, data, {
      ...geometry,
      chartArea: {
        roundedCorners: false,
        fill: { color: generated.SchemeColor.accent2, transparency: 25 },
        border: { color: '112233', pt: 2 },
      },
      plotArea: {
        fill: { color: '445566', transparency: 50 },
        border: { color: generated.SchemeColor.accent1, pt: 1 },
      },
    });
    slide.addChart(generated.ChartType.bar!, data, {
      ...geometry,
      chartArea: {
        roundedCorners: true,
        fill: { type: 'none' },
        border: { type: 'none', color: '778899', pt: 1.5 },
      },
      plotArea: {
        fill: { type: 'none' },
        border: { type: 'dash', color: 'AABBCC', pt: 2.5 },
      },
    });
    slide.addChart(generated.ChartType.bar!, data, {
      ...geometry,
      plotArea: {
        fill: { color: '654321', transparency: 25 },
        border: { color: '654321', pt: 1 },
      },
      fill: 'ABCDEF',
      border: { color: '123456', pt: 3 },
    });
    slide.addChart(generated.ChartType.bar!, data, geometry);
    slide.addChart(generated.ChartType.bar!, data, {
      ...geometry,
      chartArea: { fill: { type: 'none', color: 'C0FFEE' } },
      plotArea: { fill: { type: 'none', color: 'C0FFEE' } },
    });

    const imported = await importPptxGenJS(generated);
    const charts = imported.slides[0]!.shapes.filter(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    );
    expect(charts).toHaveLength(5);
    const [canonical, ignoredTypes, aliases, defaults, coloredNone] = charts as [
      ChartModel, ChartModel, ChartModel, ChartModel, ChartModel,
    ];

    expect(canonical.definition?.options).toMatchObject({
      chartArea: {
        fill: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 25,
        },
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '112233' },
          width: 2,
          dash: 'solid',
        },
      },
      plotArea: {
        fill: {
          kind: 'solid',
          color: { kind: 'srgb', value: '445566' },
          transparency: 50,
        },
        line: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent1' },
          width: 1,
          dash: 'solid',
        },
      },
    });
    expect(canonical.definition?.options.roundedCorners).toBeUndefined();
    expect(canonical.xml).toContain('<c:roundedCorners val="0"/>');

    expect(ignoredTypes.definition?.options).toMatchObject({
      roundedCorners: true,
      chartArea: {
        fill: { kind: 'none' },
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '778899' },
          width: 1.5,
          dash: 'solid',
        },
      },
      plotArea: {
        fill: { kind: 'none' },
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: 'AABBCC' },
          width: 2.5,
          dash: 'solid',
        },
      },
    });
    expect(ignoredTypes.xml).toContain('<c:roundedCorners val="1"/>');

    expect(aliases.definition?.options.plotArea).toMatchObject({
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'ABCDEF' },
        transparency: 25,
      },
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: '123456' },
        width: 3,
        dash: 'solid',
      },
    });
    expect(aliases.definition?.options.plotArea?.fill).not.toMatchObject({
      color: { kind: 'srgb', value: '654321' },
    });
    expect(defaults.definition?.options).toMatchObject({
      roundedCorners: true,
      plotArea: { fill: { kind: 'none' }, line: { kind: 'none' } },
    });
    expect(defaults.definition?.options.chartArea).toBeUndefined();
    expect(coloredNone.definition?.options.chartArea?.fill).toBeUndefined();
    expect(coloredNone.definition?.options.plotArea?.fill).toBeUndefined();
    expect(coloredNone.xml).not.toContain('C0FFEE');
  });

  it('isolates inherited inert IChartOpts text and top-level gridline declarations from chart output', async () => {
    const bulletNumberTypes = [
      'alphaLcParenBoth',
      'alphaLcParenR',
      'alphaLcPeriod',
      'alphaUcParenBoth',
      'alphaUcParenR',
      'alphaUcPeriod',
      'arabicParenBoth',
      'arabicParenR',
      'arabicPeriod',
      'arabicPlain',
      'romanLcParenBoth',
      'romanLcParenR',
      'romanLcPeriod',
      'romanUcParenBoth',
      'romanUcParenR',
      'romanUcPeriod',
    ] as const;
    const underlineStyles = [
      'dash',
      'dashHeavy',
      'dashLong',
      'dashLongHeavy',
      'dbl',
      'dotDash',
      'dotDashHeave',
      'dotDotDash',
      'dotDotDashHeavy',
      'dotted',
      'dottedHeavy',
      'heavy',
      'none',
      'sng',
      'wavy',
      'wavyDbl',
      'wavyHeavy',
    ] as const;
    const alignments = ['left', 'center', 'right', 'justify'] as const;
    const verticalAlignments = ['top', 'middle', 'bottom'] as const;
    const lineCaps = ['flat', 'round', 'square'] as const;
    const lineStyles = ['solid', 'dash', 'dot', 'none'] as const;
    const tabAlignments = ['l', 'r', 'ctr', 'dec'] as const;
    const textDirections = ['horz', 'vert', 'vert270', 'wordArtVert'] as const;
    const generated = new PptxGenJS();
    const slide = generated.addSlide();
    const geometry = { x: 1, y: 1, w: 6, h: 4 } as const;
    const data = () => [{ name: 'Revenue', labels: ['Q1', 'Q2'], values: [10, 20] }];
    const inertVariants: PptxGenJSPublicChartOptions[] = Array.from(
      { length: 18 },
      (_, index) => ({
        ...geometry,
        align: alignments[index % alignments.length]!,
        bold: index % 2 === 0,
        breakLine: index % 2 !== 0,
        bullet: index === 0
          ? true
          : index === 1
            ? false
            : {
                type: index % 2 === 0 ? 'bullet' : 'number',
                characterCode: '25BA',
                indent: 12,
                numberType: bulletNumberTypes[index - 2]!,
                numberStartAt: 3,
                code: '2022',
                marginPt: 4,
                startAt: 2,
                style: 'arabicPeriod',
              },
        cap: lineCaps[index % lineCaps.length]!,
        color: generated.SchemeColor.accent1,
        fontFace: 'Aptos',
        fontSize: 18,
        highlight: 'FFFF00',
        italic: index % 2 !== 0,
        size: 2,
        softBreakBefore: index % 2 === 0,
        style: lineStyles[index % lineStyles.length]!,
        tabStops: [{
          position: 1.5,
          alignment: tabAlignments[index % tabAlignments.length]!,
        }],
        textDirection: textDirections[index % textDirections.length]!,
        transparency: 25,
        underline: {
          style: underlineStyles[index % underlineStyles.length]!,
          color: '112233',
        },
        valign: verticalAlignments[index % verticalAlignments.length]!,
      }),
    );
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => undefined);

    try {
      slide.addChart(generated.ChartType.bar!, data(), geometry);
      for (const options of inertVariants) {
        slide.addChart(generated.ChartType.bar!, data(), options);
      }
      slide.addChart(generated.ChartType.bar!, data(), {
        ...geometry,
        showTitle: true,
        title: 'Observable title',
      });

      const imported = await importPptxGenJS(generated);
      const charts = imported.slides[0]!.shapes.filter(
        (shape): shape is ChartModel => shape instanceof ChartModel,
      );
      expect(charts).toHaveLength(20);
      const [baseline, ...variantsAndControl] = charts;
      const control = variantsAndControl.pop();
      for (const variant of variantsAndControl) expect(variant.xml).toBe(baseline!.xml);
      expect(control!.xml).not.toBe(baseline!.xml);
      expect(control!.xml).toContain('Observable title');
      expect(warning).not.toHaveBeenCalled();
    } finally {
      warning.mockRestore();
    }
  });

  it('projects and edits representative PptxGenJS chart options semantically', async () => {
    const generated = new PptxGenJS();
    const slide = generated.addSlide();
    slide.addChart(generated.ChartType.bar!, [{
      name: 'Revenue', labels: ['Q1', 'Q2'], values: [10, 20],
    }], {
      altText: 'Quarterly revenue chart',
      objectName: 'Revenue chart frame',
      x: 0.5,
      y: 0.5,
      w: 5,
      h: 3,
      lang: 'zh-CN',
      displayBlanksAs: 'zero',
      showTitle: true,
      title: 'Revenue',
      titleBold: true,
      titleColor: '112233',
      titleFontFace: 'Aptos Display',
      titleFontSize: 20,
      titleRotate: 30,
      titlePos: { x: 0.2, y: 0.1 },
      showLegend: true,
      legendPos: 'tr',
      legendColor: '445566',
      legendFontFace: 'Aptos',
      legendFontSize: 10,
      chartArea: {
        roundedCorners: false,
        fill: { color: 'F0F0F0' },
        border: { color: '111111', pt: 2 },
      },
      plotArea: {
        fill: { color: 'FFFFFF' },
        border: { color: '999999', pt: 1 },
      },
      chartColors: ['4472C4'],
      chartColorsOpacity: 75,
      showValue: true,
      showLabel: true,
      showSerName: true,
      showLeaderLines: true,
      dataLabelPosition: 'inEnd',
      dataLabelColor: '223344',
      dataLabelFontBold: true,
      dataLabelFontFace: 'Aptos',
      dataLabelFontSize: 9,
      dataLabelFormatCode: '0.0%',
      catAxisHidden: true,
      catAxisLabelColor: '334455',
      catAxisLabelFontFace: 'Aptos Narrow',
      catAxisLabelFontSize: 8,
      catAxisLabelRotate: -45,
      catAxisMajorTickMark: 'inside',
      catAxisLineColor: '556677',
      catAxisLineShow: true,
      catAxisLineSize: 1,
      catGridLine: { color: 'CCCCCC', size: 0.5, style: 'dash' },
      valAxisMinVal: 0,
      valAxisMaxVal: 100,
      valAxisMajorUnit: 20,
      valAxisLogScaleBase: 10,
      valAxisLabelFormatCode: '#,##0',
      valAxisMajorTickMark: 'outside',
      valGridLine: { color: 'DDDDDD', size: 0.5, style: 'dot' },
      showDataTable: true,
      showDataTableHorzBorder: false,
      showDataTableVertBorder: true,
      showDataTableOutline: false,
      showDataTableKeys: false,
      dataTableFontSize: 9,
      dataTableFormatCode: '#,##0.00',
      barDir: 'bar',
      barGrouping: 'stacked',
      barGapWidthPct: 25,
      barOverlapPct: -25,
    });
    slide.addChart(generated.ChartType.line!, [{
      name: 'Trend', labels: ['Q1', 'Q2'], values: [11, 21],
    }], {
      x: 0.5,
      y: 0.5,
      w: 5,
      h: 3,
      chartColors: ['70AD47'],
      lineSmooth: true,
      lineDataSymbol: 'diamond',
      lineDataSymbolSize: 8,
      lineDataSymbolLineColor: 'FF0000',
      lineDataSymbolLineSize: 1.5,
      lineSize: 3,
      lineDash: 'dash',
    });
    slide.addChart(generated.ChartType.doughnut!, [{
      name: 'Share', labels: ['A', 'B'], values: [60, 40],
    }], { x: 0.5, y: 0.5, w: 5, h: 3, holeSize: 65, firstSliceAng: 45 });
    slide.addChart(generated.ChartType.radar!, [{
      name: 'Score', labels: ['A', 'B'], values: [60, 40],
    }], { x: 0.5, y: 0.5, w: 5, h: 3, radarStyle: 'filled' });
    slide.addChart(generated.ChartType.bar3d!, [{
      name: '3D', labels: ['A', 'B'], values: [1, 2],
    }], {
      x: 0.5,
      y: 0.5,
      w: 5,
      h: 3,
      barGapDepthPct: 25,
      v3DRotX: -30,
      v3DRotY: 120,
      v3DRAngAx: true,
      v3DPerspective: 20,
    });

    const imported = await importPptxGenJS(generated);
    const charts = imported.slides[0]!.shapes.filter(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    );
    expect(charts).toHaveLength(5);
    const [bar, line, doughnut, radar, bar3D] = charts as [
      ChartModel, ChartModel, ChartModel, ChartModel, ChartModel,
    ];

    expect(bar).toMatchObject({
      name: 'Revenue chart frame',
      altText: 'Quarterly revenue chart',
      transform: {
        x: inches(0.5),
        y: inches(0.5),
        width: inches(5),
        height: inches(3),
      },
    });
    expect(bar.definition).toMatchObject({
      groups: [{
        type: 'bar',
        options: {
          direction: 'bar',
          grouping: 'stacked',
          gapWidth: 25,
          overlap: 100,
          dataLabels: {
            showValue: true,
            showSeriesName: true,
            showLeaderLines: true,
            position: 'insideEnd',
            numberFormat: '0.0%',
            face: 'Aptos',
            size: 9,
            bold: true,
            color: { kind: 'srgb', value: '223344' },
          },
          series: [{
            fill: {
              kind: 'solid',
              color: { kind: 'srgb', value: '4472C4' },
              transparency: 25,
            },
          }],
        },
      }],
      options: {
        displayBlanksAs: 'span',
        title: {
          text: 'Revenue',
          rotation: 30,
          face: 'Aptos Display',
          size: 20,
          bold: true,
          color: { kind: 'srgb', value: '112233' },
        },
        legend: {
          position: 'topRight',
          face: 'Aptos',
          size: 10,
          color: { kind: 'srgb', value: '445566' },
        },
        chartArea: {
          fill: { kind: 'solid', color: { kind: 'srgb', value: 'F0F0F0' } },
        },
        plotArea: {
          fill: { kind: 'solid', color: { kind: 'srgb', value: 'FFFFFF' } },
        },
        categoryAxis: {
          visible: false,
          labelRotation: -45,
        },
        valueAxis: {
          minimum: 0,
          maximum: 100,
          majorUnit: 20,
          logarithmicBase: 10,
          numberFormat: '#,##0',
        },
        dataTable: {
          showHorizontalBorder: false,
          showOutline: false,
          showLegendKeys: false,
          numberFormat: '#,##0.00',
          size: 9,
        },
      },
    });
    expect(bar.definition?.options.title?.position).toBeDefined();
    expect(line.definition).toMatchObject({
      groups: [{
        type: 'line',
        options: {
          smooth: true,
          series: [{ marker: { shape: 'diamond', size: 8 } }],
        },
      }],
    });
    expect(doughnut.definition).toMatchObject({
      groups: [{ options: { holeSize: 65, firstSliceAngle: 45 } }],
    });
    expect(radar.definition).toMatchObject({ groups: [{ options: { style: 'filled' } }] });
    expect(bar3D.definition).toMatchObject({
      groups: [{ options: { gapDepth: 25 } }],
      options: { rotationX: -30, rotationY: 120, perspective: 20 },
    });

    const workbookBefore = imported.opcPackage.requirePart(bar.workbookPartUri!).bytes.slice();
    await bar.replaceDefinition({
      groups: bar.definition!.groups,
      options: {
        ...bar.definition!.options,
        title: { ...bar.definition!.options.title, text: 'Revenue edited' },
      },
    });
    expect(bar.definition?.options.title?.text).toBe('Revenue edited');
    expect(imported.opcPackage.requirePart(bar.workbookPartUri!).bytes).toEqual(workbookBefore);
  });

  it('promotes uniform PptxGenJS point labels and preserves custom scatter labels', async () => {
    const generated = new PptxGenJS();
    const cases = [
      {
        name: 'pie-label-true-percent-false',
        type: generated.ChartType.pie!,
        options: { dataLabelPosition: 'bestFit', showLabel: true, showPercent: false },
      },
      {
        name: 'pie-label-false-percent-true',
        type: generated.ChartType.pie!,
        options: { dataLabelPosition: 'bestFit', showLabel: false, showPercent: true },
      },
      {
        name: 'pie-label-false-percent-false',
        type: generated.ChartType.pie!,
        options: { dataLabelPosition: 'bestFit', showLabel: false, showPercent: false },
      },
      {
        name: 'doughnut-label-true-percent-false',
        type: generated.ChartType.doughnut!,
        options: { dataLabelPosition: 'bestFit', showLabel: true, showPercent: false },
      },
    ] as const;
    for (const entry of cases) {
      generated.addSlide().addChart(entry.type, [{
        name: entry.name, labels: ['A', 'B', 'C'], values: [10, 20, 30],
      }], { x: 0.5, y: 0.5, w: 5, h: 3, ...entry.options });
    }
    generated.addSlide().addChart(generated.ChartType.scatter!, [
      { name: 'X', values: [1, 2, 3] },
      { name: 'Custom', labels: ['A', 'B', 'C'], values: [10, 20, 30] },
    ], {
      x: 0.5,
      y: 0.5,
      w: 5,
      h: 3,
      dataLabelPosition: 't',
      dataLabelFormatScatter: 'custom',
      showLabel: true,
    });

    const imported = await importPptxGenJS(generated);
    const charts = imported.slides.map((slide) => slide.shapes.find(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    )!);
    expect(charts).toHaveLength(5);
    expect(charts.map((chart) => chart.definition?.groups[0]?.options?.dataLabels)).toEqual([
      {
        face: 'Arial',
        size: 12,
        color: { kind: 'srgb', value: '000000' },
        showCategoryName: true,
        position: 'bestFit',
        numberFormat: 'General',
      },
      {
        face: 'Arial',
        size: 12,
        color: { kind: 'srgb', value: '000000' },
        showPercent: true,
        position: 'bestFit',
        numberFormat: '0%',
      },
      {
        face: 'Arial',
        size: 12,
        color: { kind: 'srgb', value: '000000' },
        position: 'bestFit',
        numberFormat: 'General',
      },
      {
        face: 'Arial',
        size: 12,
        color: { kind: 'srgb', value: '000000' },
        showCategoryName: true,
        numberFormat: 'General',
      },
      {
        face: 'Arial',
        size: 12,
        color: { kind: 'srgb', value: '000000' },
        position: 'top',
        numberFormat: 'General',
      },
    ]);

    const pie = charts[0]!;
    const pieLabels = pie.xml.match(/<c:dLbls>[\s\S]*?<\/c:dLbls>/gu)![0]!;
    const pieWorkbookBefore = imported.opcPackage.requirePart(pie.workbookPartUri!).bytes.slice();
    const pieDefinition = pie.definition!;
    await pie.replaceDefinition({
      groups: pieDefinition.groups,
      options: { ...pieDefinition.options, title: { text: 'Unrelated pie title' } },
    });
    expect(pie.xml).toContain(pieLabels);
    expect(imported.opcPackage.requirePart(pie.workbookPartUri!).bytes).toEqual(pieWorkbookBefore);

    const titledPie = pie.definition!;
    const titledPieGroup = titledPie.groups[0];
    if (titledPieGroup?.type !== 'pie') throw new Error('Expected imported pie chart');
    await pie.replaceDefinition({
      groups: [{
        type: 'pie',
        series: titledPieGroup.series,
        options: { dataLabels: { position: 'center', showPercent: true } },
      }],
      options: titledPie.options,
    });
    expect(pie.xml).not.toContain(pieLabels);
    expect(pie.definition?.groups[0]?.options?.dataLabels).toEqual({
      showPercent: true,
      position: 'center',
    });
    expect(imported.opcPackage.requirePart(pie.workbookPartUri!).bytes)
      .toEqual(pieWorkbookBefore);

    const scatter = charts[4]!;
    const scatterLabels = scatter.xml.match(/<c:dLbls>[\s\S]*?<\/c:dLbls>/gu)!;
    expect(scatterLabels.join('')).toContain('<c:tx>');
    expect(scatterLabels.join('')).toContain('<c:spPr>');
    expect(scatterLabels.join('')).toContain('<c:extLst>');
    const scatterDefinition = scatter.definition!;
    await scatter.replaceDefinition({
      groups: scatterDefinition.groups,
      options: { ...scatterDefinition.options, title: { text: 'Unrelated scatter title' } },
    });
    expect(scatter.xml.match(/<c:dLbls>[\s\S]*?<\/c:dLbls>/gu)).toEqual(scatterLabels);

    const reopened = await PptxDocument.open(await imported.write());
    expect(reopened.slides.map((slide) => slide.shapes.find(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    )?.definition?.groups[0]?.options?.dataLabels)).toEqual(
      charts.map((chart) => chart.definition?.groups[0]?.options?.dataLabels),
    );
  });

  it('compares chart axis line gridline visibility label and tick semantics', async () => {
    const labelPositions = ['high', 'low', 'nextTo', 'none'] as const;
    const tickMarks = ['cross', 'inside', 'none', 'outside'] as const;
    const lineStyles = ['solid', 'dash', 'dot', 'solid'] as const;
    const gridStyles = ['solid', 'dash', 'dot', 'none'] as const;
    const rotations = [-90, -45, 45, 90] as const;
    const generated = new PptxGenJS();
    const generatedSlide = generated.addSlide();
    for (const index of labelPositions.keys()) {
      const reverse = labelPositions.length - index - 1;
      generatedSlide.addChart(generated.ChartType.bar!, [{
        name: `Series ${index + 1}`,
        labels: ['Q1', 'Q2'],
        values: [10 + index, 20 + index],
      }], {
        x: 0.5,
        y: 0.5,
        w: 5,
        h: 3,
        catAxisHidden: index % 2 === 0,
        catAxisLabelPos: labelPositions[index]!,
        catAxisLabelRotate: rotations[index]!,
        catAxisMajorTickMark: tickMarks[index]!,
        catAxisMinorTickMark: tickMarks[reverse]!,
        catAxisLineColor: '112233',
        catAxisLineShow: index !== 3,
        catAxisLineSize: index + 1,
        catAxisLineStyle: lineStyles[index]!,
        catGridLine: {
          color: '445566',
          size: 0.5 + index * 0.25,
          style: gridStyles[index]!,
        },
        valAxisHidden: index % 2 !== 0,
        valAxisLabelPos: labelPositions[reverse]!,
        valAxisLabelRotate: rotations[reverse]!,
        valAxisMajorTickMark: tickMarks[reverse]!,
        valAxisMinorTickMark: tickMarks[index]!,
        valAxisLineColor: '778899',
        valAxisLineShow: index !== 0,
        valAxisLineSize: index + 1.5,
        valAxisLineStyle: lineStyles[reverse]!,
        valGridLine: {
          color: 'AABBCC',
          size: 0.75 + index * 0.25,
          style: gridStyles[reverse]!,
        },
      });
    }

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedCharts = imported.slides[0]!.shapes.filter(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    );
    expect(importedCharts).toHaveLength(4);
    const axisFragment = (xml: string, name: 'catAx' | 'valAx') => {
      const fragment = xml.match(new RegExp(`<c:${name}>[\\s\\S]*?</c:${name}>`, 'u'))?.[0];
      expect(fragment).toBeDefined();
      return fragment!;
    };
    const axisShapeFragment = (xml: string) => {
      const fragment = xml.match(
        /<c:tickLblPos[^>]*\/>\s*<c:spPr>([\s\S]*?)<\/c:spPr>/u,
      )?.[1];
      expect(fragment).toBeDefined();
      return fragment!;
    };
    const gridLineFragment = (xml: string) =>
      xml.match(/<c:majorGridlines>([\s\S]*?)<\/c:majorGridlines>/u)?.[1];

    for (const [index, chart] of importedCharts.entries()) {
      const reverse = labelPositions.length - index - 1;
      const categoryXml = axisFragment(chart.xml, 'catAx');
      const valueXml = axisFragment(chart.xml, 'valAx');
      expect(categoryXml).toContain(`<c:delete val="${index % 2 === 0 ? 1 : 0}"/>`);
      expect(valueXml).toContain(`<c:delete val="${index % 2 !== 0 ? 1 : 0}"/>`);
      expect(categoryXml).toContain(`<c:tickLblPos val="${labelPositions[index]}"/>`);
      expect(valueXml).toContain(`<c:tickLblPos val="${labelPositions[reverse]}"/>`);
      expect(categoryXml).toContain(`<c:majorTickMark val="${tickMarks[index]}"/>`);
      expect(categoryXml).toContain(`<c:minorTickMark val="${tickMarks[reverse]}"/>`);
      expect(valueXml).toContain(`<c:majorTickMark val="${tickMarks[reverse]}"/>`);
      expect(valueXml).toContain(`<c:minorTickMark val="${tickMarks[index]}"/>`);
      expect(categoryXml).toContain(`rot="${rotations[index]! * 60_000}"`);
      expect(valueXml).toContain(`rot="${rotations[reverse]! * 60_000}"`);

      const categoryShape = axisShapeFragment(categoryXml);
      const valueShape = axisShapeFragment(valueXml);
      expect(categoryShape).toContain(`<a:ln w="${(index + 1) * 12_700}"`);
      expect(valueShape).toContain(`<a:ln w="${(index + 1.5) * 12_700}"`);
      expect(categoryShape).toContain(`<a:prstDash val="${lineStyles[index]}"/>`);
      expect(valueShape).toContain(`<a:prstDash val="${lineStyles[reverse]}"/>`);
      expect(categoryShape).toContain(index === 3 ? '<a:noFill/>' : 'val="112233"');
      expect(valueShape).toContain(index === 0 ? '<a:noFill/>' : 'val="778899"');

      const categoryGrid = gridLineFragment(categoryXml);
      const valueGrid = gridLineFragment(valueXml);
      if (gridStyles[index] === 'none') expect(categoryGrid).toBeUndefined();
      else {
        expect(categoryGrid).toContain(
          `<a:ln w="${(0.5 + index * 0.25) * 12_700}"`,
        );
        expect(categoryGrid).toContain('val="445566"');
        expect(categoryGrid).toContain(`<a:prstDash val="${gridStyles[index]}"/>`);
      }
      if (gridStyles[reverse] === 'none') expect(valueGrid).toBeUndefined();
      else {
        expect(valueGrid).toContain(
          `<a:ln w="${(0.75 + index * 0.25) * 12_700}"`,
        );
        expect(valueGrid).toContain('val="AABBCC"');
        expect(valueGrid).toContain(`<a:prstDash val="${gridStyles[reverse]}"/>`);
      }

      const expectedCategoryPosition = labelPositions[index] === 'nextTo'
        ? undefined
        : labelPositions[index];
      const expectedValuePosition = labelPositions[reverse] === 'nextTo'
        ? undefined
        : labelPositions[reverse];
      expect(chart.definition?.options.categoryAxis).toMatchObject({
        ...(index % 2 === 0 ? { visible: false } : {}),
        ...(expectedCategoryPosition ? { labelPosition: expectedCategoryPosition } : {}),
        labelRotation: rotations[index],
      });
      expect(chart.definition?.options.categoryAxis?.labelPosition)
        .toBe(expectedCategoryPosition);
      expect(chart.definition?.options.valueAxis).toMatchObject({
        ...(index % 2 !== 0 ? { visible: false } : {}),
        ...(expectedValuePosition ? { labelPosition: expectedValuePosition } : {}),
        labelRotation: rotations[reverse],
      });
      expect(chart.definition?.options.valueAxis?.labelPosition).toBe(expectedValuePosition);
    }

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    for (const index of labelPositions.keys()) {
      const reverse = labelPositions.length - index - 1;
      const nativeChart = await nativeSlide.addChart('bar', [{
        name: `Native ${index + 1}`,
        categories: ['Q1', 'Q2'],
        values: [10 + index, 20 + index],
      }]);
      const categoryGridStyle = gridStyles[index]!;
      const valueGridStyle = gridStyles[reverse]!;
      const categoryLineStyle = lineStyles[index]!;
      const valueLineStyle = lineStyles[reverse]!;
      await nativeChart.replaceDefinition({
        groups: nativeChart.definition!.groups,
        options: {
          categoryAxis: {
            visible: index % 2 !== 0,
            labelPosition: labelPositions[index]!,
            labelRotation: rotations[index]!,
            line: index === 3
              ? { kind: 'none' }
              : {
                  kind: 'line',
                  color: { kind: 'srgb', value: '112233' },
                  width: index + 1,
                  dash: categoryLineStyle === 'dot' ? 'sysDot' : categoryLineStyle,
                },
            majorGridLine: categoryGridStyle === 'none'
              ? { kind: 'none' }
              : {
                  kind: 'line',
                  color: { kind: 'srgb', value: '445566' },
                  width: 0.5 + index * 0.25,
                  dash: categoryGridStyle === 'dot' ? 'sysDot' : categoryGridStyle,
                },
            majorTickMark: tickMarks[index]!,
            minorTickMark: tickMarks[reverse]!,
          },
          valueAxis: {
            visible: index % 2 === 0,
            labelPosition: labelPositions[reverse]!,
            labelRotation: rotations[reverse]!,
            line: index === 0
              ? { kind: 'none' }
              : {
                  kind: 'line',
                  color: { kind: 'srgb', value: '778899' },
                  width: index + 1.5,
                  dash: valueLineStyle === 'dot' ? 'sysDot' : valueLineStyle,
                },
            majorGridLine: valueGridStyle === 'none'
              ? { kind: 'none' }
              : {
                  kind: 'line',
                  color: { kind: 'srgb', value: 'AABBCC' },
                  width: 0.75 + index * 0.25,
                  dash: valueGridStyle === 'dot' ? 'sysDot' : valueGridStyle,
                },
            majorTickMark: tickMarks[reverse]!,
            minorTickMark: tickMarks[index]!,
          },
        },
      });
    }
    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedNativeCharts = reopenedNative.slides[0]!.shapes.filter(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    );
    expect(reopenedNativeCharts).toHaveLength(4);
    const tickXml = {
      cross: 'cross', inside: 'in', none: 'none', outside: 'out',
    } as const;
    for (const [index, chart] of reopenedNativeCharts.entries()) {
      const reverse = labelPositions.length - index - 1;
      const categoryXml = axisFragment(chart.xml, 'catAx');
      const valueXml = axisFragment(chart.xml, 'valAx');
      const categoryLineStyle = lineStyles[index]!;
      const valueLineStyle = lineStyles[reverse]!;
      const categoryGridStyle = gridStyles[index]!;
      const valueGridStyle = gridStyles[reverse]!;
      const expectedCategoryLine = index === 3
        ? { kind: 'none' }
        : {
            kind: 'line',
            color: { kind: 'srgb', value: '112233' },
            width: index + 1,
            dash: categoryLineStyle === 'dot' ? 'sysDot' : categoryLineStyle,
          };
      const expectedValueLine = index === 0
        ? { kind: 'none' }
        : {
            kind: 'line',
            color: { kind: 'srgb', value: '778899' },
            width: index + 1.5,
            dash: valueLineStyle === 'dot' ? 'sysDot' : valueLineStyle,
          };
      const expectedCategoryGrid = categoryGridStyle === 'none'
        ? { kind: 'none' }
        : {
            kind: 'line',
            color: { kind: 'srgb', value: '445566' },
            width: 0.5 + index * 0.25,
            dash: categoryGridStyle === 'dot' ? 'sysDot' : categoryGridStyle,
          };
      const expectedValueGrid = valueGridStyle === 'none'
        ? { kind: 'none' }
        : {
            kind: 'line',
            color: { kind: 'srgb', value: 'AABBCC' },
            width: 0.75 + index * 0.25,
            dash: valueGridStyle === 'dot' ? 'sysDot' : valueGridStyle,
          };
      expect(chart.definition?.options.categoryAxis?.line).toEqual(expectedCategoryLine);
      expect(chart.definition?.options.valueAxis?.line).toEqual(expectedValueLine);
      expect(chart.definition?.options.categoryAxis?.majorGridLine)
        .toEqual(expectedCategoryGrid);
      expect(chart.definition?.options.valueAxis?.majorGridLine).toEqual(expectedValueGrid);
      expect(categoryXml).toContain(`<c:tickLblPos val="${labelPositions[index]}"/>`);
      expect(valueXml).toContain(`<c:tickLblPos val="${labelPositions[reverse]}"/>`);
      expect(categoryXml).toContain(
        `<c:majorTickMark val="${tickXml[tickMarks[index]!]}"/>`,
      );
      expect(categoryXml).toContain(
        `<c:minorTickMark val="${tickXml[tickMarks[reverse]!]}"/>`,
      );
      expect(valueXml).toContain(
        `<c:majorTickMark val="${tickXml[tickMarks[reverse]!]}"/>`,
      );
      expect(valueXml).toContain(
        `<c:minorTickMark val="${tickXml[tickMarks[index]!]}"/>`,
      );
      expect(categoryXml).toContain(`rot="${rotations[index]! * 60_000}"`);
      expect(valueXml).toContain(`rot="${rotations[reverse]! * 60_000}"`);
      expect(chart.xml).not.toContain('val="inside"');
      expect(chart.xml).not.toContain('val="outside"');
      expect(chart.xml).not.toContain('<a:prstDash val="dot"/>');
      if (categoryLineStyle === 'dot' || categoryGridStyle === 'dot') {
        expect(categoryXml).toContain('<a:prstDash val="sysDot"/>');
      }
      if (valueLineStyle === 'dot' || valueGridStyle === 'dot') {
        expect(valueXml).toContain('<a:prstDash val="sysDot"/>');
      }
      expect(chart.definition?.options.categoryAxis?.labelPosition)
        .toBe(labelPositions[index] === 'nextTo' ? undefined : labelPositions[index]);
      expect(chart.definition?.options.valueAxis?.labelPosition)
        .toBe(labelPositions[reverse] === 'nextTo' ? undefined : labelPositions[reverse]);
    }
  });

  it('compares advanced date value display-unit and series-axis semantics', async () => {
    const displayUnits = [
      'billions',
      'hundredMillions',
      'hundredThousands',
      'hundreds',
      'millions',
      'tenMillions',
      'tenThousands',
      'thousands',
      'trillions',
    ] as const;
    const generated = new PptxGenJS();
    const slide = generated.addSlide();
    for (const [index, valAxisDisplayUnit] of displayUnits.entries()) {
      const options: PptxGenJSPublicChartOptions = {
        x: 0.5,
        y: 0.5,
        w: 5,
        h: 3,
        valAxisDisplayUnit,
        valAxisDisplayUnitLabel: index !== 1,
        ...(index === 0 ? {
          catLabelFormatCode: 'yyyy-mm-dd',
          catAxisBaseTimeUnit: 'days',
          catAxisMajorTimeUnit: 'months',
          catAxisMinorTimeUnit: 'years',
          catAxisMajorUnit: 2,
          catAxisMinorUnit: 3,
          catAxisLabelFrequency: '4',
          catAxisMultiLevelLabels: true,
          valAxisCrossesAt: 2,
          catAxisCrossesAt: 4,
          serAxisHidden: true,
          serAxisLabelFrequency: '5',
          serAxisLabelPos: 'low',
          serAxisMajorUnit: 2,
          serAxisMinorUnit: 3,
          serAxisOrientation: 'maxMin',
          serLabelFormatCode: 'yyyy',
          serAxisLabelColor: 'AABBCC',
          serAxisLabelFontBold: true,
          serAxisLabelFontItalic: true,
          serAxisLabelFontFace: 'Aptos',
          serAxisLabelFontSize: 11,
          serAxisLineColor: '112233',
          serAxisLineShow: true,
          serGridLine: { color: '445566', size: 1.5, style: 'dash' },
          showSerAxisTitle: true,
          serAxisTitle: 'Series title',
          serAxisTitleColor: '778899',
          serAxisTitleFontFace: 'Aptos Display',
          serAxisTitleFontSize: 15,
          serAxisTitleRotate: 25,
        } : {}),
      };
      slide.addChart(index === 0 ? generated.ChartType.bar3d! : generated.ChartType.bar!, [{
        name: `Axis ${valAxisDisplayUnit}`,
        labels: index === 0 ? ['46023', '46054'] : ['Q1', 'Q2'],
        values: [10, 20],
      }], options);
    }

    const imported = await openPptxGenJSPublicOutput(generated);
    const charts = imported.slides[0]!.shapes.filter(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    );
    expect(charts).toHaveLength(displayUnits.length);
    for (const [index, chart] of charts.entries()) {
      expect(chart.definition?.options.valueAxis?.displayUnit).toBe(displayUnits[index]);
      expect(chart.definition?.options.valueAxis?.displayUnitLabel)
        .toBe(index === 1 ? undefined : true);
    }
    expect(charts[0]!.definition?.options.categoryAxis).toMatchObject({
      kind: 'date',
      numberFormat: 'yyyy-mm-dd',
      crossesAt: 2,
      majorUnit: 2,
      minorUnit: 3,
      baseTimeUnit: 'days',
      majorTimeUnit: 'months',
      minorTimeUnit: 'years',
      labelFrequency: 4,
      multiLevelLabels: true,
    });
    expect(charts[0]!.definition?.options.valueAxis).toMatchObject({
      crossesAt: 4,
      displayUnit: 'billions',
      displayUnitLabel: true,
    });
    expect(charts[0]!.definition?.options.seriesAxis).toMatchObject({
      visible: false,
      numberFormat: 'yyyy',
      orientation: 'maxMin',
      labelPosition: 'low',
      labelFrequency: 5,
      majorUnit: 2,
      minorUnit: 3,
      face: 'Aptos',
      size: 11,
      bold: true,
      italic: true,
      color: { kind: 'srgb', value: 'AABBCC' },
      line: {
        kind: 'line', color: { kind: 'srgb', value: '112233' }, width: 1, dash: 'solid',
      },
      majorGridLine: {
        kind: 'line', color: { kind: 'srgb', value: '445566' }, width: 1.5, dash: 'dash',
      },
      title: {
        text: 'Series title', rotation: 25, face: 'Aptos Display', size: 15,
        color: { kind: 'srgb', value: '778899' },
      },
    });
    expect(imported.diagnostics.filter(({ code }) => code.startsWith('CHART_'))).toEqual([]);
  });

  it('locks PptxGenJS chart axis position and series-axis runtime defects', async () => {
    const generated = new PptxGenJS();
    const slide = generated.addSlide();
    const data = [{ name: 'Axis', labels: ['A', 'B'], values: [1, 2] }];
    slide.addChart(generated.ChartType.bar!, data, {});
    for (const axisPos of ['b', 'l', 'r', 't'] as const) {
      slide.addChart(generated.ChartType.bar!, data, { axisPos });
    }
    for (const serAxisLabelPos of ['low', 'high', 'nextTo', 'none'] as const) {
      slide.addChart(generated.ChartType.bar3d!, data, {
        serAxisLabelPos,
        ...(serAxisLabelPos === 'low' ? {
          serLabelFormatCode: 'yyyy',
          serAxisBaseTimeUnit: 'days',
          serAxisMajorTimeUnit: 'months',
          serAxisMinorTimeUnit: 'years',
        } : {}),
      });
    }

    const warnings: string[] = [];
    const warning = vi.spyOn(console, 'warn').mockImplementation((message) => {
      warnings.push(String(message));
    });
    try {
      const imported = await openPptxGenJSPublicOutput(generated);
      const charts = imported.slides[0]!.shapes.filter(
        (shape): shape is ChartModel => shape instanceof ChartModel,
      );
      expect(charts).toHaveLength(9);
      const axisPair = (xml: string) => ['catAx', 'valAx'].map((name) =>
        xml.match(new RegExp(`<c:${name}>[\\s\\S]*?</c:${name}>`, 'u'))?.[0]);
      const baseline = axisPair(charts[0]!.xml);
      for (const chart of charts.slice(1, 5)) expect(axisPair(chart.xml)).toEqual(baseline);

      const seriesAxes = charts.slice(5).map((chart) =>
        chart.xml.match(/<c:serAx>[\s\S]*?<\/c:serAx>/u)?.[0]);
      expect(seriesAxes.every((xml) => xml?.includes('<c:tickLblPos val="low"/>'))).toBe(true);
      expect(seriesAxes[0]).not.toMatch(/<c:(?:base|major|minor)TimeUnit\b/u);
      expect(warnings).toEqual([
        '"serAxisBaseTimeUnit" must be one of: \'days\',\'months\',\'years\' !',
        '"serAxisMajorTimeUnit" must be one of: \'days\',\'months\',\'years\' !',
        '"serAxisMinorTimeUnit" must be one of: \'days\',\'months\',\'years\' !',
      ]);
    } finally {
      warning.mockRestore();
    }
  });

  it('matches valid PptxGenJS public audio and video media output semantically', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'pptxgenjs-media-'));
    const audioPath = join(directory, 'path-audio.mp3');
    const pathBytes = Uint8Array.of(9, 10, 11, 12);
    await writeFile(audioPath, pathBytes);
    try {
      const generated = new PptxGenJS();
      expect(generated.version).toBe('4.0.1');
      const generatedSlide = generated.addSlide();
      const publicAudioOptions: PptxGenJSMediaOptions = {
        type: 'audio',
        data: 'data:audio/mpeg;base64,AQIDBA==',
        cover: PNG_DATA_URI,
        extn: 'mp3',
        objectName: 'Audio & narration',
        x: 1,
        y: 1,
        w: 2,
        h: 1,
      };
      generatedSlide.addMedia(publicAudioOptions);
      generatedSlide.addMedia({
        type: 'video',
        data: 'data:video/mp4;base64,BQYHCA==',
        cover: PNG_DATA_URI,
        extn: 'mp4',
        objectName: 'Video overview',
        x: 2,
        y: 2,
        w: 3,
        h: 2,
      });
      generatedSlide.addMedia({
        type: 'audio',
        path: audioPath,
        cover: PNG_DATA_URI,
        objectName: 'Path audio',
        x: 3,
        y: 3,
        w: 2,
        h: 1,
      });
      generatedSlide.addMedia({
        type: 'audio',
        path: audioPath,
        cover: PNG_DATA_URI,
        objectName: 'Duplicate path audio',
        x: 4,
        y: 4,
        w: 2,
        h: 1,
      });

      const imported = await openPptxGenJSPublicOutput(generated);
      const native = PptxDocument.create();
      native.addSlide();
      await native.addAudio(0, 'data:audio/mpeg;base64,AQIDBA==', {
        name: 'Audio & narration',
        poster: PNG_DATA_URI,
        x: inches(1),
        y: inches(1),
        width: inches(2),
        height: inches(1),
      });
      await native.addVideo(0, 'data:video/mp4;base64,BQYHCA==', {
        name: 'Video overview',
        poster: PNG_DATA_URI,
        x: inches(2),
        y: inches(2),
        width: inches(3),
        height: inches(2),
      });
      await native.addAudio(0, audioPath, {
        name: 'Path audio',
        poster: PNG_DATA_URI,
        x: inches(3),
        y: inches(3),
        width: inches(2),
        height: inches(1),
      });
      await native.addAudio(0, audioPath, {
        name: 'Duplicate path audio',
        poster: PNG_DATA_URI,
        x: inches(4),
        y: inches(4),
        width: inches(2),
        height: inches(1),
      });

      const importedStates = embeddedMediaStates(imported, 0);
      const nativeStates = embeddedMediaStates(native, 0);
      expect(importedStates).toHaveLength(4);
      expect(nativeStates).toHaveLength(4);
      expect(importedStates.map(comparableEmbeddedMediaState))
        .toEqual(nativeStates.map(comparableEmbeddedMediaState));
      expect(importedStates.map(({ nameXml }) => nameXml)).toEqual([
        'Audio &amp; narration',
        'Video overview',
        'Path audio',
        'Duplicate path audio',
      ]);
      expect(importedStates[2]!.mediaPartUri).toBe(importedStates[3]!.mediaPartUri);
      expect(nativeStates[2]!.mediaPartUri).toBe(nativeStates[3]!.mediaPartUri);
      for (const index of [0, 1, 2]) {
        expect(importedStates[index]!.relationshipRoles)
          .toEqual(nativeStates[index]!.relationshipRoles);
      }
      expect(importedStates[0]!.relationshipRoles).toEqual({
        kind: {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio',
          targetMode: 'Internal',
          resolvesToMedia: true,
        },
        media: {
          type: OFFICE_MEDIA_RELATIONSHIP,
          targetMode: 'Internal',
          resolvesToMedia: true,
        },
        poster: {
          type: IMAGE_RELATIONSHIP,
          targetMode: 'Internal',
          resolvesToPoster: true,
        },
      });
      expect(importedStates[1]!.relationshipRoles.kind.type)
        .toBe('http://schemas.openxmlformats.org/officeDocument/2006/relationships/video');
      expect(importedStates[3]!.relationshipRoles.kind.type).toBe(OFFICE_MEDIA_RELATIONSHIP);
      expect(nativeStates[3]!.relationshipRoles.kind.type)
        .toBe('http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio');
      expect(importedStates.map(({ fileElement }) => fileElement))
        .toEqual(['videoFile', 'videoFile', 'videoFile', 'videoFile']);
      expect(nativeStates.map(({ fileElement }) => fileElement))
        .toEqual(['audioFile', 'videoFile', 'audioFile', 'audioFile']);
      expect(importedStates.map(({ mediaContentType }) => mediaContentType))
        .toEqual(['audio/mp3', 'video/mp4', 'audio/mp3', 'audio/mp3']);
      expect(nativeStates.map(({ mediaContentType }) => mediaContentType))
        .toEqual(['audio/mpeg', 'video/mp4', 'audio/mpeg', 'audio/mpeg']);

      await imported.write();
      await native.write();
      expect(imported.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
      expect(native.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
      const reopenedImported = await PptxDocument.open(await imported.write());
      const reopenedNative = await PptxDocument.open(await native.write());
      expect(embeddedMediaStates(reopenedImported, 0).map(comparableEmbeddedMediaState))
        .toEqual(importedStates.map(comparableEmbeddedMediaState));
      expect(embeddedMediaStates(reopenedNative, 0).map(comparableEmbeddedMediaState))
        .toEqual(nativeStates.map(comparableEmbeddedMediaState));
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  }, 20_000);

  it('locks MediaProps source, type, metadata, and geometry against PptxGenJS 4.0.1', async () => {
    const link = 'https://www.youtube.com/embed/dQw4w9WgXcQ?rel=0&controls=0';
    const generated = new PptxGenJS();
    generated.addSlide().addMedia({
      type: 'online',
      link,
      data: 'data:video/mp4;base64,AQID',
      path: 'ignored-online-source.mp4',
      extn: 'mp4',
      cover: PNG_DATA_URI,
      objectName: 'Online & overview',
      x: 1,
      y: 1.5,
      w: 4,
      h: 2.25,
    });
    const imported = await openPptxGenJSPublicOutput(generated);

    const native = PptxDocument.create();
    native.addSlide();
    await native.addVideo(0, link, {
      poster: PNG_DATA_URI,
      name: 'Online & overview',
      x: inches(1),
      y: inches(1.5),
      width: inches(4),
      height: inches(2.25),
    });

    const importedState = externalVideoState(imported, 0);
    const nativeState = externalVideoState(native, 0);
    const importedRelationshipsXml = slideRelationshipsXml(imported, 0);
    const nativeRelationshipsXml = slideRelationshipsXml(native, 0);
    expect(importedRelationshipsXml).toContain(
      `Target="${link}" TargetMode="External"`,
    );
    expect(importedRelationshipsXml).not.toContain('rel=0&amp;controls=0');
    expect(nativeRelationshipsXml.match(/rel=0&amp;controls=0/g)).toHaveLength(1);
    expect(nativeRelationshipsXml).not.toContain(`Target="${link}"`);
    expect(nativeRelationshipsXml).not.toContain('&amp;amp;');
    expect(comparableExternalVideoState(importedState))
      .toEqual(comparableExternalVideoState(nativeState));
    expect(importedState).toMatchObject({
      settings: {},
      clickAction: false,
      aspectLocked: false,
      playbackExtension: false,
      timing: false,
    });
    expect(nativeState).toMatchObject({
      settings: {
        play: 'click',
        loop: false,
        hideWhenStopped: false,
        volume: 1,
      },
      clickAction: true,
      aspectLocked: true,
      playbackExtension: true,
      timing: true,
    });

    const reopenedImported = await PptxDocument.open(await imported.write());
    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImportedState = externalVideoState(reopenedImported, 0);
    const reopenedNativeState = externalVideoState(reopenedNative, 0);
    expect(reopenedImportedState).toMatchObject({
      externalUrl: link,
      relationships: { video: { target: link } },
    });
    expect(reopenedNativeState).toMatchObject({
      externalUrl: link,
      relationships: { video: { target: link } },
    });
    expect(comparableExternalVideoState(reopenedImportedState))
      .toEqual(comparableExternalVideoState(importedState));
    expect(comparableExternalVideoState(reopenedNativeState))
      .toEqual(comparableExternalVideoState(nativeState));

    const publicRejected = new PptxGenJS().addSlide();
    expect(() => publicRejected.addMedia({
      type: 'online',
      cover: PNG_DATA_URI,
    } as PptxGenJSMediaOptions)).toThrow(/online videos require.*link/i);
    expect(() => publicRejected.addMedia({
      type: 'online',
      link: 'javascript:alert(1)',
      cover: PNG_DATA_URI,
    })).not.toThrow();

    const before = packageState(native);
    for (const source of ['javascript:alert(1)', 'ftp://example.com/video.mp4']) {
      await expect(native.addVideo(0, source, { poster: PNG_DATA_URI })).rejects.toThrow();
      expect(packageState(native)).toEqual(before);
    }
  }, 20_000);

  it('locks PptxGenJS media defects while native creation remains strict', async () => {
    const generated = new PptxGenJS();
    const generatedSlide = generated.addSlide();
    generatedSlide.addMedia({
      type: 'audio',
      data: 'data:audio/mpeg;base64,AQIDBA==',
      cover: PNG_DATA_URI,
      extn: 'mp3',
      objectName: 'PptxGenJS audio',
    });
    const imported = await openPptxGenJSPublicOutput(generated);
    const native = PptxDocument.create();
    native.addSlide();
    await native.addAudio(0, 'data:audio/mpeg;base64,AQIDBA==', {
      name: 'Native audio',
      poster: PNG_DATA_URI,
    });
    const importedAudio = embeddedMediaStates(imported, 0)[0]!;
    const nativeAudio = embeddedMediaStates(native, 0)[0]!;

    expect(importedAudio.fileElement).toBe('videoFile');
    expect(nativeAudio.fileElement).toBe('audioFile');
    expect(importedAudio.mediaContentType).toBe('audio/mp3');
    expect(nativeAudio.mediaContentType).toBe('audio/mpeg');
    expect(importedAudio.mediaExtension).toBe('.mp3');
    expect(nativeAudio.mediaExtension).toBe('.mp3');

    const rejected = new PptxGenJS().addSlide();
    expect(() => rejected.addMedia({ type: 'audio' } as PptxGenJSMediaOptions))
      .toThrow(/data.*path.*required/i);
    expect(() => rejected.addMedia({
      type: 'audio',
      data: 'not-base64',
      cover: PNG_DATA_URI,
    })).toThrow(/base64 header/i);
    expect(() => rejected.addMedia({
      type: 'audio',
      data: 'data:audio/mpeg;base64,AQ==',
      cover: 'not-base64',
    })).toThrow(/cover.*base64 header/i);

    const coerced = new PptxGenJS().addSlide();
    expect(() => coerced.addMedia({
      type: 'audio',
      data: 'data:audio/mpeg;base64,A===',
      cover: PNG_DATA_URI,
      extn: 'mp3',
    })).not.toThrow();
    expect(() => coerced.addMedia({
      type: 'audio',
      data: 'data:audio/mpeg;base64,AQ==',
      cover: PNG_DATA_URI,
      extn: 'wav',
    })).not.toThrow();
    expect(() => coerced.addMedia({
      type: 'audio',
      data: 'data:audio/mpeg;base64,AQ==',
      cover: PNG_DATA_URI,
      extn: 'mp3',
      x: Number.NaN,
    })).not.toThrow();
    expect(() => coerced.addMedia({
      type: 'music' as never,
      data: 'data:audio/mpeg;base64,AQ==',
      cover: PNG_DATA_URI,
      extn: 'mp3',
    })).not.toThrow();
    expect(() => (coerced.addMedia as (options: unknown) => void)(null)).toThrow();

    const before = packageState(native);
    const invalidNativeCalls: Array<() => Promise<unknown>> = [
      () => native.addAudio(0, ''),
      () => native.addAudio(0, new Uint8Array()),
      () => native.addAudio(0, 'data:audio/mpeg;base64,A===', {}),
      () => native.addAudio(0, Uint8Array.of(1), {
        poster: 'data:image/png;base64,A===',
      }),
      () => native.addAudio(0, 'data:video/mp4;base64,AQ==', {}),
      () => native.addAudio(0, Uint8Array.of(1), {
        contentType: 'audio/mpeg',
        fileName: 'voice.wav',
      }),
      () => native.addAudio(0, Uint8Array.of(1), { name: 1 } as never),
    ];
    for (const call of invalidNativeCalls) {
      await expect(call()).rejects.toThrow();
      expect(packageState(native)).toEqual(before);
    }
  }, 20_000);

  it('opens PptxGenJS media and appends canonical native media without changing it', async () => {
    const generated = new PptxGenJS();
    const generatedSlide = generated.addSlide();
    generatedSlide.addMedia({
      type: 'audio',
      data: 'data:audio/mpeg;base64,AQID',
      cover: PNG_DATA_URI,
      extn: 'mp3',
      objectName: 'Existing audio',
      x: 1,
      y: 1,
      w: 2,
      h: 1,
    });
    generatedSlide.addMedia({
      type: 'video',
      data: 'data:video/mp4;base64,BAUG',
      cover: PNG_DATA_URI,
      extn: 'mp4',
      objectName: 'Existing video',
      x: 2,
      y: 2,
      w: 3,
      h: 2,
    });
    const imported = await importPptxGenJS(generated);
    expect(imported.media(0)).toHaveLength(2);
    const existingStates = embeddedMediaStates(imported, 0);
    const existingPictures = existingStates.map(({ shapeId }) => pictureXml(imported, 0, shapeId));
    const existingParts = new Map(existingStates.flatMap(({ mediaPartUri, posterPartUri }) => [
      [mediaPartUri, [...imported.opcPackage.requirePart(mediaPartUri).bytes]] as const,
      [posterPartUri, [...imported.opcPackage.requirePart(posterPartUri).bytes]] as const,
    ]));
    const existingRelationships = imported.slides[0]!.relationships.map(
      (relationship) => ({ ...relationship }),
    );

    const created = await imported.addAudio(0, Uint8Array.of(17, 18, 19), {
      name: 'Strict & native',
      altText: 'Canonical addition',
      contentType: 'audio/mpeg',
      poster: 'data:image/jpeg;base64,FRYX',
      x: inches(3),
      y: inches(3),
      width: inches(2),
      height: inches(1),
    });
    const reopened = await PptxDocument.open(await imported.write());
    expect(reopened.media(0)).toHaveLength(3);
    const reopenedStates = embeddedMediaStates(reopened, 0);
    expect(reopenedStates).toHaveLength(3);
    for (const [index, source] of existingStates.entries()) {
      expect(pictureXml(reopened, 0, source.shapeId)).toBe(existingPictures[index]);
    }
    for (const [uri, bytes] of existingParts) {
      expect([...reopened.opcPackage.requirePart(uri).bytes]).toEqual(bytes);
    }
    for (const relationship of existingRelationships) {
      expect(reopened.slides[0]!.relationships.find(({ id }) => id === relationship.id))
        .toEqual(relationship);
    }

    const strict = reopenedStates[2]!;
    expect(strict).toMatchObject({
      fileElement: 'audioFile',
      nameXml: 'Strict &amp; native',
      mediaPartUri: created.mediaPartUri,
      mediaContentType: 'audio/mpeg',
      mediaExtension: '.mp3',
      posterPartUri: created.posterPartUri,
      posterContentType: 'image/jpeg',
      posterExtension: '.jpg',
      relationshipRoles: {
        kind: {
          type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio',
          targetMode: 'Internal',
          resolvesToMedia: true,
        },
        media: {
          type: OFFICE_MEDIA_RELATIONSHIP,
          targetMode: 'Internal',
          resolvesToMedia: true,
        },
        poster: {
          type: IMAGE_RELATIONSHIP,
          targetMode: 'Internal',
          resolvesToPoster: true,
        },
      },
    });
    expect(pictureXml(reopened, 0, strict.shapeId)).toContain('descr="Canonical addition"');
    await reopened.write();
    expect(reopened.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
  }, 20_000);

  it('edits, canonicalizes, isolates, and removes PptxGenJS legacy media safely', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'pptxgenjs-media-lifecycle-'));
    const audioPath = join(directory, 'duplicate.mp3');
    const videoPath = join(directory, 'legacy-video.mp4');
    await writeFile(audioPath, Uint8Array.of(9, 10, 11));
    await writeFile(videoPath, Uint8Array.of(4, 5, 6));
    try {
      const generated = new PptxGenJS();
      const generatedSlide = generated.addSlide();
      generatedSlide.addMedia({
        type: 'audio',
        data: 'data:audio/mpeg;base64,AQID',
        cover: PNG_DATA_URI,
        extn: 'mp3',
        objectName: 'Legacy audio',
        x: 1,
        y: 1,
        w: 2,
        h: 1,
      });
      generatedSlide.addMedia({
        type: 'video',
        path: videoPath,
        cover: PNG_DATA_URI,
        objectName: 'Legacy video',
        x: 2,
        y: 2,
        w: 3,
        h: 2,
      });
      generatedSlide.addMedia({
        type: 'audio',
        path: audioPath,
        cover: PNG_DATA_URI,
        objectName: 'Duplicate audio A',
      });
      generatedSlide.addMedia({
        type: 'audio',
        path: audioPath,
        cover: PNG_DATA_URI,
        objectName: 'Duplicate audio B',
      });

      const imported = await openPptxGenJSPublicOutput(generated);
      const [audio, video, duplicateA, duplicateB] = imported.media(0);
      expect(audio).toBeInstanceOf(MediaModel);
      expect(imported.slides[0]!.shapes[0]).toBe(audio);
      expect(imported.slides[0]!.media[0]).toBe(audio);
      expect(duplicateA!.mediaPartUri).toBe(duplicateB!.mediaPartUri);
      expect(slideXml(imported, 0)).not.toContain('<p:timing>');
      expect(slideXml(imported, 0)).not.toContain('<px:playback');
      const noOpState = packageState(imported);
      await imported.write();
      expect(packageState(imported)).toEqual(noOpState);
      const duplicateBefore = pictureXml(imported, 0, duplicateB!.shapeId);

      audio!.name = 'Legacy audio edited';
      audio!.altText = 'Edited without canonicalization';
      audio!.settings = { play: 'auto', loop: true, volume: 0.5 };
      audio!.setTransform({
        x: inches(2),
        y: inches(1),
        width: inches(3),
        height: inches(1),
      });
      const preserved = embeddedMediaStates(imported, 0)[0]!;
      expect(preserved.fileElement).toBe('videoFile');
      expect(preserved.mediaContentType).toBe('audio/mp3');
      expect(preserved.relationshipRoles.kind.type)
        .toBe('http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio');
      expect(slideXml(imported, 0).match(/<p:cMediaNode\b/g)).toHaveLength(1);
      expect(slideXml(imported, 0)).toContain(`spid="${audio!.shapeId}"`);
      expect(slideXml(imported, 0)).toContain('cmd="playFrom(0.0)"');

      expect(await audio!.replaceSource(Uint8Array.of(12, 13), {
        contentType: 'audio/mpeg',
      })).toBe(audio);
      const canonical = embeddedMediaStates(imported, 0)[0]!;
      expect(canonical.fileElement).toBe('audioFile');
      expect(canonical.mediaContentType).toBe('audio/mpeg');
      expect(canonical.relationshipRoles.kind.type)
        .toBe('http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio');
      expect(pictureXml(imported, 0, duplicateB!.shapeId)).toBe(duplicateBefore);

      const sharedPoster = duplicateA!.posterPartUri;
      expect(await duplicateB!.replacePoster(Uint8Array.of(14), {
        contentType: 'image/gif',
      })).toBe(duplicateB);
      expect(duplicateA!.posterPartUri).toBe(sharedPoster);
      expect(duplicateB!.posterPartUri).not.toBe(sharedPoster);
      const sharedMedia = duplicateB!.mediaPartUri!;
      duplicateA!.remove();
      expect(imported.media(0)).not.toContain(duplicateA);
      expect(duplicateB!.mediaPartUri).toBe(sharedMedia);
      expect(imported.opcPackage.hasPart(sharedMedia)).toBe(true);
      expect(video!.name).toBe('Legacy video');

      const sourceBeforeDuplicate = slideXml(imported, 0);
      const duplicateSlide = imported.duplicateSlide(0);
      const isolated = duplicateSlide.media[0]!;
      expect(isolated.mediaPartUri).toBe(audio!.mediaPartUri);
      isolated.settings = {
        play: 'click',
        loop: false,
        hideWhenStopped: true,
        volume: 0.25,
      };
      expect(slideXml(imported, 0)).toBe(sourceBeforeDuplicate);
      await isolated.replaceSource(Uint8Array.of(15), { contentType: 'audio/wav' });
      await isolated.replacePoster(Uint8Array.of(16), { contentType: 'image/jpeg' });
      expect(isolated.mediaPartUri).not.toBe(audio!.mediaPartUri);
      expect(isolated.posterPartUri).not.toBe(audio!.posterPartUri);

      const reopened = await PptxDocument.open(await imported.write());
      expect(reopened.media(0).map(({ name }) => name)).toEqual([
        'Legacy audio edited',
        'Legacy video',
        'Duplicate audio B',
      ]);
      expect(reopened.media(1)).toHaveLength(3);
      expect(reopened.media(0)[0]!.settings).toEqual({
        play: 'auto',
        loop: true,
        hideWhenStopped: false,
        volume: 0.5,
      });
      expect(reopened.media(1)[0]!.settings).toEqual({
        play: 'click',
        loop: false,
        hideWhenStopped: true,
        volume: 0.25,
      });
      expect(slideXml(reopened, 0).match(/<p:cMediaNode\b/g)).toHaveLength(1);
      expect(slideXml(reopened, 1).match(/<p:cMediaNode\b/g)).toHaveLength(1);
      expect(embeddedMediaStates(reopened, 0)[2]).toMatchObject({
        fileElement: 'videoFile',
        mediaContentType: 'audio/mp3',
      });
      expect(reopened.opcPackage.requirePart(reopened.media(0)[2]!.posterPartUri!)).toMatchObject({
        contentType: 'image/gif',
        bytes: Uint8Array.of(14),
      });
      reopened.media(0)[0]!.settings = {
        play: 'click',
        hideWhenStopped: true,
        volume: 0.25,
      };
      const editedAgain = await PptxDocument.open(await reopened.write());
      expect(editedAgain.media(0)[0]!.settings).toEqual({
        play: 'click',
        loop: false,
        hideWhenStopped: true,
        volume: 0.25,
      });
      await editedAgain.write({ mode: 'permissive' });
      expect(editedAgain.diagnostics.filter(({ code }) => code.startsWith('MEDIA_TIMING_')))
        .toEqual([]);
      expect(reopened.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  }, 20_000);

  it('matches PptxGenJS embedded raster image public output semantically', async () => {
    const fixtures = [
      { data: PNG_DATA_URI, contentType: 'image/png' as const, name: 'Raster PNG' },
      { data: JPEG_DATA_URI, contentType: 'image/jpeg' as const, name: 'Raster JPEG' },
      { data: GIF_DATA_URI, contentType: 'image/gif' as const, name: 'Raster GIF' },
    ];
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    for (const fixture of fixtures) {
      generatedSlide.addImage({
        data: fixture.data,
        x: 1,
        y: 2,
        w: 3,
        h: 2,
        rotate: 45,
        flipH: true,
        flipV: false,
        objectName: fixture.name,
        altText: 'Raster alt',
      });
    }

    const imported = await importPptxGenJS(generated);
    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    for (const fixture of fixtures) {
      nativeSlide.addImage(
        new Uint8Array(Buffer.from(fixture.data.split(',')[1]!, 'base64')),
        {
          contentType: fixture.contentType,
          name: fixture.name,
          altText: 'Raster alt',
          x: inches(1),
          y: inches(2),
          width: inches(3),
          height: inches(2),
          rotation: degrees(45),
          flipHorizontal: true,
          flipVertical: false,
        },
      );
    }

    const importedImages = imported.slides[0]!.shapes.filter(
      (shape): shape is ImageModel => shape instanceof ImageModel,
    );
    const nativeImages = nativeSlide.shapes.filter(
      (shape): shape is ImageModel => shape instanceof ImageModel,
    );
    expect(imported.slides[0]!.shapes.map(({ kind }) => kind))
      .toEqual(nativeSlide.shapes.map(({ kind }) => kind));
    expect(importedImages).toHaveLength(fixtures.length);
    expect(nativeImages).toHaveLength(fixtures.length);
    for (const [index, importedImage] of importedImages.entries()) {
      expect(embeddedRasterState(imported, 0, importedImage))
        .toEqual(embeddedRasterState(native, 0, nativeImages[index]!));
    }
  }, 20_000);

  it('matches PptxGenJS image identity and visual effects through create edit reopen', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    generatedSlide.addImage({
      data: PNG_DATA_URI,
      x: 1,
      y: 1,
      w: 2,
      h: 2,
      objectName: 'Data Image & Name',
      altText: 'Data alt & <accessible>',
      rounding: true,
      transparency: 25,
      shadow: {
        type: 'outer',
        color: '123456',
        opacity: 0.5,
        blur: 3,
        angle: 30,
        offset: 2,
        rotateWithShape: true,
      },
    });
    generatedSlide.addImage({
      data: PNG_DATA_URI,
      x: 4,
      y: 1,
      w: 2,
      h: 2,
      objectName: 'Baseline image',
      altText: 'Baseline alt',
      rounding: false,
      transparency: 0,
    });

    const imported = await importPptxGenJS(generated);
    const [styled, baseline] = imported.slides[0]!.shapes as readonly ImageModel[];
    expect({
      name: styled!.name,
      altText: styled!.altText,
      rounding: styled!.rounding,
      transparency: styled!.transparency,
      shadow: styled!.shadow,
    }).toEqual({
      name: 'Data Image & Name',
      altText: 'Data alt & <accessible>',
      rounding: true,
      transparency: 25,
      shadow: {
        kind: 'outer',
        color: { kind: 'srgb', value: '123456' },
        opacity: 0.5,
        blur: 3,
        angle: 30,
        distance: 2,
        rotateWithShape: false,
      },
    });
    expect({
      rounding: baseline!.rounding,
      transparency: baseline!.transparency,
      shadow: baseline!.shadow,
    }).toEqual({ rounding: false, transparency: 0, shadow: undefined });
    const importedXml = pictureXml(imported, 0, styled!.id);
    expect(importedXml).toContain('name="Data Image &amp; Name"');
    expect(importedXml).toContain('descr="Data alt &amp; &lt;accessible&gt;"');
    expect(importedXml).toContain('<a:prstGeom prst="ellipse">');
    expect(importedXml).toContain('<a:alphaModFix amt="75000"/>');
    expect(importedXml).toContain('rotWithShape="0"');

    const sourcePartUri = styled!.sourcePartUri;
    const relationshipSnapshot = imported.slides[0]!.relationships;
    const mediaBytes = imported.opcPackage.requirePart(sourcePartUri!).bytes.slice();
    const beforeNoop = packageState(imported);
    styled!.name = styled!.name;
    styled!.altText = styled!.altText;
    styled!.rounding = styled!.rounding!;
    styled!.transparency = styled!.transparency!;
    styled!.shadow = styled!.shadow;
    expect(packageState(imported)).toEqual(beforeNoop);

    styled!.name = 'Edited image';
    styled!.altText = 'Edited alt';
    styled!.rounding = false;
    styled!.transparency = 40;
    styled!.shadow = {
      kind: 'outer',
      color: { kind: 'srgb', value: '654321' },
      opacity: 0.75,
      blur: 4,
      angle: 45,
      distance: 3,
      rotateWithShape: true,
    };
    expect(imported.slides[0]!.relationships).toEqual(relationshipSnapshot);
    expect(imported.opcPackage.requirePart(sourcePartUri!).bytes).toEqual(mediaBytes);

    const duplicateSlide = imported.duplicateSlide(0);
    const duplicate = duplicateSlide.shapes[0] as ImageModel;
    duplicate.altText = undefined;
    duplicate.transparency = 0;
    duplicate.shadow = undefined;
    expect(styled!.altText).toBe('Edited alt');
    expect(styled!.transparency).toBe(40);
    expect(styled!.shadow?.kind).toBe('outer');

    const reopened = await PptxDocument.open(await imported.write());
    const reopenedStyled = reopened.slides[0]!.shapes[0] as ImageModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as ImageModel;
    expect({
      name: reopenedStyled.name,
      altText: reopenedStyled.altText,
      rounding: reopenedStyled.rounding,
      transparency: reopenedStyled.transparency,
      shadow: reopenedStyled.shadow,
    }).toEqual({
      name: 'Edited image',
      altText: 'Edited alt',
      rounding: false,
      transparency: 40,
      shadow: {
        kind: 'outer',
        color: { kind: 'srgb', value: '654321' },
        opacity: 0.75,
        blur: 4,
        angle: 45,
        distance: 3,
        rotateWithShape: true,
      },
    });
    expect(reopenedDuplicate.altText).toBeUndefined();
    expect(reopenedDuplicate.transparency).toBe(0);
    expect(reopenedDuplicate.shadow).toBeUndefined();
    expect(reopenedStyled.sourcePartUri).toBe(sourcePartUri);
    expect(reopenedDuplicate.sourcePartUri).toBe(sourcePartUri);
    expect(reopened.opcPackage.requirePart(sourcePartUri!).bytes).toEqual(mediaBytes);
  }, 20_000);

  it('matches PptxGenJS path and data images through the document source loader', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'pptxgenjs-raster-source-'));
    try {
      const fixtures: readonly {
        readonly data: string;
        readonly path?: string;
        readonly name: string;
      }[] = [
        { data: PNG_DATA_URI, path: join(directory, 'source.png'), name: 'Path PNG' },
        { data: JPEG_DATA_URI, name: 'Data JPEG' },
        { data: GIF_DATA_URI, path: join(directory, 'source.gif'), name: 'Path GIF' },
      ];
      for (const fixture of fixtures) {
        if (fixture.path) {
          await writeFile(fixture.path, Buffer.from(fixture.data.split(',')[1]!, 'base64'));
        }
      }

      const generated = new PptxGenJS();
      expect(generated.version).toBe('4.0.1');
      const generatedSlide = generated.addSlide();
      for (const fixture of fixtures) {
        generatedSlide.addImage({
          ...(fixture.path ? { path: fixture.path } : { data: fixture.data }),
          x: 1,
          y: 2,
          w: 3,
          h: 2,
          rotate: 45,
          flipH: true,
          flipV: false,
          objectName: fixture.name,
          altText: 'Portable source',
        });
      }

      const imported = await importPptxGenJS(generated);
      const native = PptxDocument.create();
      native.addSlide();
      for (const fixture of fixtures) {
        await native.addImage(0, fixture.path ?? fixture.data, {
          name: fixture.name,
          altText: 'Portable source',
          x: inches(1),
          y: inches(2),
          width: inches(3),
          height: inches(2),
          rotation: degrees(45),
          flipHorizontal: true,
          flipVertical: false,
        });
      }

      const importedImages = imported.slides[0]!.shapes as readonly ImageModel[];
      const nativeImages = native.slides[0]!.shapes as readonly ImageModel[];
      expect(importedImages).toHaveLength(fixtures.length);
      expect(nativeImages).toHaveLength(fixtures.length);
      for (const [index, importedImage] of importedImages.entries()) {
        expect(embeddedRasterState(imported, 0, importedImage))
          .toEqual(embeddedRasterState(native, 0, nativeImages[index]!));
      }
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  }, 20_000);

  it('matches PptxGenJS contain, cover, and crop sizing final state', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'pptxgenjs-raster-sizing-'));
    try {
      const landscape = pngHeader(1600, 900);
      const portrait = pngHeader(900, 1600);
      const square = pngHeader(1000, 1000);
      const landscapePath = join(directory, 'landscape.png');
      await writeFile(landscapePath, landscape);
      const cases = [
        {
          name: 'Contain landscape',
          generatedSource: { path: landscapePath },
          nativeSource: landscapePath,
          outer: { w: 16, h: 9 },
          generatedSizing: { type: 'contain', w: 4, h: 3 },
          nativeSizing: { type: 'contain', width: inches(4), height: inches(3) },
          raw: { left: 0, top: -16_667, right: 0, bottom: -16_667 },
        },
        {
          name: 'Contain portrait',
          generatedSource: { data: pngDataUri(portrait) },
          nativeSource: pngDataUri(portrait),
          outer: { w: 9, h: 16 },
          generatedSizing: { type: 'contain', w: 4, h: 3 },
          nativeSizing: { type: 'contain', width: inches(4), height: inches(3) },
          raw: { left: -68_519, top: 0, right: -68_519, bottom: 0 },
        },
        {
          name: 'Cover landscape',
          generatedSource: { data: pngDataUri(landscape) },
          nativeSource: pngDataUri(landscape),
          outer: { w: 16, h: 9 },
          generatedSizing: { type: 'cover', w: 4, h: 3 },
          nativeSizing: { type: 'cover', width: inches(4), height: inches(3) },
          raw: { left: 12_500, top: 0, right: 12_500, bottom: 0 },
        },
        {
          name: 'Cover portrait',
          generatedSource: { data: pngDataUri(portrait) },
          nativeSource: pngDataUri(portrait),
          outer: { w: 9, h: 16 },
          generatedSizing: { type: 'cover', w: 4, h: 3 },
          nativeSizing: { type: 'cover', width: inches(4), height: inches(3) },
          raw: { left: 0, top: 28_906, right: 0, bottom: 28_906 },
        },
        {
          name: 'Equal square',
          generatedSource: { data: pngDataUri(square) },
          nativeSource: pngDataUri(square),
          outer: { w: 10, h: 10 },
          generatedSizing: { type: 'cover', w: 4, h: 4 },
          nativeSizing: { type: 'cover', width: inches(4), height: inches(4) },
          raw: { left: 0, top: 0, right: 0, bottom: 0 },
        },
        {
          name: 'Crop landscape center',
          generatedSource: { path: landscapePath },
          nativeSource: landscapePath,
          outer: { w: 16, h: 9 },
          generatedSizing: { type: 'crop', x: 4, y: 2.25, w: 8, h: 4.5 },
          nativeSizing: {
            type: 'crop',
            width: inches(8),
            height: inches(4.5),
            source: { x: 400, y: 225, width: 800, height: 450 },
          },
          raw: { left: 25_000, top: 25_000, right: 25_000, bottom: 25_000 },
        },
      ] as const;

      const generated = new PptxGenJS();
      expect(generated.version).toBe('4.0.1');
      const generatedSlide = generated.addSlide();
      const native = PptxDocument.create();
      native.addSlide();
      for (const entry of cases) {
        generatedSlide.addImage({
          ...entry.generatedSource,
          x: 1,
          y: 0.5,
          w: entry.outer.w,
          h: entry.outer.h,
          sizing: entry.generatedSizing,
          rotate: 45,
          flipH: true,
          flipV: true,
          objectName: entry.name,
          altText: `${entry.name} alt`,
        });
        await native.addImage(0, entry.nativeSource, {
          x: inches(1),
          y: inches(0.5),
          sizing: entry.nativeSizing,
          rotation: degrees(45),
          flipHorizontal: true,
          flipVertical: true,
          name: entry.name,
          altText: `${entry.name} alt`,
        });
      }

      const imported = await importPptxGenJS(generated);
      const importedImages = imported.slides[0]!.shapes as readonly ImageModel[];
      const nativeImages = native.slides[0]!.shapes as readonly ImageModel[];
      expect(importedImages).toHaveLength(cases.length);
      expect(nativeImages).toHaveLength(cases.length);
      expect(importedImages.map(({ name }) => name)).toEqual(cases.map(({ name }) => name));
      for (const [index, entry] of cases.entries()) {
        expect(embeddedRasterSizingState(imported, 0, importedImages[index]!))
          .toEqual(embeddedRasterSizingState(native, 0, nativeImages[index]!));
        expect(embeddedRasterSizingState(native, 0, nativeImages[index]!).sourceRectangleRaw)
          .toEqual(entry.raw);
      }
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  }, 20_000);

  it('matches PptxGenJS SVG data/path pictures, sizing, and OOXML roles', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'pptxgenjs-svg-conformance-'));
    try {
      const svgPath = join(directory, 'source.svg');
      await writeFile(svgPath, SVG_BYTES);
      const fallback = new Uint8Array(Buffer.from(PNG_DATA_URI.split(',')[1]!, 'base64'));
      const cases = [
        {
          name: 'Data contain SVG',
          generatedSource: { data: SVG_DATA_URI },
          nativeSource: SVG_DATA_URI,
          generatedSizing: { type: 'contain', w: 4, h: 3 },
          nativeSizing: { type: 'contain', width: inches(4), height: inches(3) },
          raw: { left: 0, top: -16_667, right: 0, bottom: -16_667 },
        },
        {
          name: 'Path cover SVG',
          generatedSource: { path: svgPath },
          nativeSource: svgPath,
          generatedSizing: { type: 'cover', w: 4, h: 3 },
          nativeSizing: { type: 'cover', width: inches(4), height: inches(3) },
          raw: { left: 12_500, top: 0, right: 12_500, bottom: 0 },
        },
        {
          name: 'Data crop SVG',
          generatedSource: { data: SVG_DATA_URI },
          nativeSource: SVG_DATA_URI,
          generatedSizing: { type: 'crop', x: 4, y: 2.25, w: 8, h: 4.5 },
          nativeSizing: {
            type: 'crop',
            width: inches(8),
            height: inches(4.5),
            source: { x: 400, y: 225, width: 800, height: 450 },
          },
          raw: { left: 25_000, top: 25_000, right: 25_000, bottom: 25_000 },
        },
      ] as const;

      const generated = new PptxGenJS();
      expect(generated.version).toBe('4.0.1');
      const generatedSlide = generated.addSlide();
      const native = PptxDocument.create();
      native.addSlide();
      for (const [index, entry] of cases.entries()) {
        const x = index + 1;
        const y = index + 0.5;
        const rotation = 15 * (index + 1);
        const flipHorizontal = index !== 1;
        const flipVertical = index !== 0;
        generatedSlide.addImage({
          ...entry.generatedSource,
          x,
          y,
          w: 16,
          h: 9,
          sizing: entry.generatedSizing,
          rotate: rotation,
          flipH: flipHorizontal,
          flipV: flipVertical,
          objectName: entry.name,
          altText: `${entry.name} alt`,
        });
        await native.addImage(0, entry.nativeSource, {
          fallback,
          x: inches(x),
          y: inches(y),
          sizing: entry.nativeSizing,
          rotation: degrees(rotation),
          flipHorizontal,
          flipVertical,
          name: entry.name,
          altText: `${entry.name} alt`,
        });
      }

      const imported = await importPptxGenJS(generated);
      const importedImages = imported.slides[0]!.shapes as readonly ImageModel[];
      const nativeImages = native.slides[0]!.shapes as readonly ImageModel[];
      expect(importedImages.map(({ name }) => name)).toEqual(cases.map(({ name }) => name));
      expect(nativeImages.map(({ name }) => name)).toEqual(cases.map(({ name }) => name));
      for (const [index, entry] of cases.entries()) {
        const importedImage = importedImages[index]!;
        const nativeImage = nativeImages[index]!;
        expect(embeddedSvgState(imported, 0, importedImage))
          .toEqual(embeddedSvgState(native, 0, nativeImage));
        expect(embeddedSvgState(native, 0, nativeImage)).toMatchObject({
          extensionUri: SVG_EXTENSION_URI,
          extensionNamespace: SVG_NAMESPACE,
          sourceRectangleRaw: entry.raw,
          fallbackContentType: 'image/png',
          svgContentType: 'image/svg+xml',
          svgBytes: [...SVG_BYTES],
          relationshipRoles: {
            fallback: {
              type: IMAGE_RELATIONSHIP,
              targetMode: 'Internal',
              resolvesToExpectedPart: true,
            },
            svg: {
              type: IMAGE_RELATIONSHIP,
              targetMode: 'Internal',
              resolvesToExpectedPart: true,
            },
          },
        });
        const nativeFallback = native.opcPackage.requirePart(nativeImage.fallbackPartUri!).bytes;
        expect(nativeFallback.byteLength).toBeGreaterThan(24);
        expect(inspectRasterImage(nativeFallback).contentType).toBe('image/png');
        const importedFallback = imported.opcPackage
          .requirePart(importedImage.fallbackPartUri!).bytes;
        if ('path' in entry.generatedSource) {
          expect(importedFallback).toEqual(SVG_BYTES);
          expect(() => inspectRasterImage(importedFallback)).toThrow();
        } else {
          expect(inspectRasterImage(importedFallback).contentType).toBe('image/png');
        }
      }
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  }, 30_000);

  it('imports and atomically edits PptxGenJS SVG pairs without disturbing neighbors', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'pptxgenjs-svg-edit-'));
    try {
      const svgPath = join(directory, 'editable.svg');
      await writeFile(svgPath, SVG_BYTES);
      const generated = new PptxGenJS();
      const generatedSlide = generated.addSlide();
      for (const [index, source] of [
        { data: SVG_DATA_URI },
        { path: svgPath },
        { data: SVG_DATA_URI },
      ].entries()) {
        generatedSlide.addImage({
          ...source,
          x: index,
          y: index,
          w: 4,
          h: 2.25,
          objectName: `Imported SVG ${index + 1}`,
          altText: `Neighbor ${index + 1}`,
        });
      }

      const imported = await openPptxGenJSPublicOutput(generated);
      const images = imported.slides[0]!.shapes as readonly ImageModel[];
      expect(images).toHaveLength(3);
      expect(images.every(({ isSvg, fallbackPartUri, svgPartUri }) =>
        isSvg && fallbackPartUri !== undefined && svgPartUri !== undefined)).toBe(true);
      const neighborParts = [images[0]!, images[2]!].map((image) => ({
        fallbackPartUri: image.fallbackPartUri!,
        fallbackBytes: new Uint8Array(
          imported.opcPackage.requirePart(image.fallbackPartUri!).bytes,
        ),
        svgPartUri: image.svgPartUri!,
        svgBytes: new Uint8Array(imported.opcPackage.requirePart(image.svgPartUri!).bytes),
      }));
      const edited = images[1]!;
      const editedFallbackPartUri = edited.fallbackPartUri!;
      const editedSvgPartUri = edited.svgPartUri!;
      const replacementSvg = new TextEncoder().encode(
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"/>',
      );
      const replacementFallback = new Uint8Array(
        Buffer.from(PNG_DATA_URI.split(',')[1]!, 'base64'),
      );

      edited.replaceSvgData(replacementSvg, replacementFallback);

      expect(edited.fallbackPartUri).toBe(editedFallbackPartUri);
      expect(edited.svgPartUri).toBe(editedSvgPartUri);
      for (const neighbor of neighborParts) {
        expect(imported.opcPackage.requirePart(neighbor.fallbackPartUri).bytes)
          .toEqual(neighbor.fallbackBytes);
        expect(imported.opcPackage.requirePart(neighbor.svgPartUri).bytes)
          .toEqual(neighbor.svgBytes);
      }

      const reopened = await PptxDocument.open(await imported.write());
      const reopenedImages = reopened.slides[0]!.shapes as readonly ImageModel[];
      expect(reopenedImages.map(({ name }) => name)).toEqual([
        'Imported SVG 1',
        'Imported SVG 2',
        'Imported SVG 3',
      ]);
      expect(reopenedImages.every(({ isSvg, fallbackPartUri, svgPartUri }) =>
        isSvg && fallbackPartUri !== undefined && svgPartUri !== undefined)).toBe(true);
      expect(reopened.opcPackage.requirePart(reopenedImages[1]!.fallbackPartUri!).bytes)
        .toEqual(replacementFallback);
      expect(reopened.opcPackage.requirePart(reopenedImages[1]!.svgPartUri!).bytes)
        .toEqual(replacementSvg);
      for (const neighbor of neighborParts) {
        expect(reopened.opcPackage.requirePart(neighbor.fallbackPartUri).bytes)
          .toEqual(neighbor.fallbackBytes);
        expect(reopened.opcPackage.requirePart(neighbor.svgPartUri).bytes)
          .toEqual(neighbor.svgBytes);
      }
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  }, 30_000);

  it('records PptxGenJS sizing fallbacks while native rejects ambiguous or unsafe state', async () => {
    const landscape = pngHeader(1600, 900);
    const generated = new PptxGenJS();
    generated.addSlide().addImage({
      data: pngDataUri(landscape),
      x: 0,
      y: 0,
      w: 16,
      h: 9,
      sizing: { type: 'cover', w: 0, h: 3 },
      objectName: 'Falsy sizing width',
    });
    const imported = await importPptxGenJS(generated);
    const importedImage = imported.slides[0]!.shapes[0] as ImageModel;
    expect(importedImage.transform).toMatchObject({ width: inches(16), height: inches(3) });
    expect(importedImage.sourceRectangle).toEqual({
      left: 0,
      top: 33.333,
      right: 0,
      bottom: 33.333,
    });

    const native = PptxDocument.create();
    native.addSlide();
    const before = packageState(native);
    const invalid: readonly {
      readonly source: string;
      readonly options: Record<string, unknown>;
    }[] = [
      {
        source: pngDataUri(landscape),
        options: {
          width: inches(16),
          height: inches(9),
          sizing: { type: 'cover', width: inches(4), height: inches(3) },
        },
      },
      {
        source: pngDataUri(landscape),
        options: { sourceRectangle: { left: 0, top: 0, right: 0, bottom: 0 } },
      },
      {
        source: pngDataUri(landscape),
        options: { sizing: { type: 'cover', width: 1.5, height: 1 } },
      },
      {
        source: pngDataUri(landscape),
        options: {
          sizing: {
            type: 'crop',
            width: 1,
            height: 1,
            source: { x: 1590, y: 0, width: 11, height: 1 },
          },
        },
      },
      {
        source: pngDataUri(pngHeader(0xffff_ffff, 1)),
        options: {
          sizing: { type: 'contain', width: 1, height: Number.MAX_SAFE_INTEGER },
        },
      },
    ];
    for (const entry of invalid) {
      await expect(native.addImage(0, entry.source, entry.options as never)).rejects.toThrow();
      expect(packageState(native)).toEqual(before);
    }
  }, 20_000);

  it('preserves PptxGenJS embedded raster image divergences while native stays strict', async () => {
    const zero = new PptxGenJS();
    zero.addSlide().addImage({
      data: PNG_DATA_URI,
      x: 0,
      y: 0,
      w: 0,
      h: 0,
      rotate: 0,
      flipH: false,
      flipV: false,
      objectName: 'Falsy transform',
    });
    const importedZero = await importPptxGenJS(zero);
    const zeroImage = importedZero.slides[0]!.shapes[0] as ImageModel;
    expect(zeroImage).toBeInstanceOf(ImageModel);
    expect(zeroImage.transform).toEqual({
      x: 0,
      y: 0,
      width: inches(1),
      height: inches(1),
      rotation: 0,
      flipHorizontal: false,
      flipVertical: false,
    });

    const rejected = new PptxGenJS();
    const rejectedSlide = rejected.addSlide();
    const error = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    let importedRejected: PptxDocument;
    try {
      rejectedSlide.addImage({ data: 'not-a-base64-data-uri' });
      rejectedSlide.addImage({ path: 42 });
      rejectedSlide.addImage({ data: 42 });
      rejectedSlide.addImage({});
      expect(error).toHaveBeenCalledTimes(4);
      importedRejected = await importPptxGenJS(rejected);
    } finally {
      error.mockRestore();
    }
    expect(importedRejected!.slides[0]!.shapes).toEqual([]);

    const quirks = new PptxGenJS();
    const quirksSlide = quirks.addSlide();
    quirksSlide.addImage({
      data: JPEG_DATA_URI.replace('image/jpeg', 'image/jpg'),
      objectName: 'JPG MIME',
    });
    quirksSlide.addImage({
      path: '/not-read/paired-source.jpeg',
      data: PNG_DATA_URI,
      objectName: 'Data and path',
    });
    const importedQuirks = await importPptxGenJS(quirks);
    const quirkImages = importedQuirks.slides[0]!.shapes as readonly ImageModel[];
    expect(quirkImages).toHaveLength(2);
    expect(quirkImages[0]).toBeInstanceOf(ImageModel);
    expect(quirkImages[0]!.sourcePartUri).toMatch(/\.jpg$/);
    expect(importedQuirks.opcPackage.requirePart(quirkImages[0]!.sourcePartUri!))
      .toMatchObject({ contentType: 'image/jpg' });
    expect(quirkImages[1]!.sourcePartUri).toMatch(/\.png$/);
    expect(importedQuirks.opcPackage.requirePart(quirkImages[1]!.sourcePartUri!).bytes)
      .toEqual(new Uint8Array(Buffer.from(PNG_DATA_URI.split(',')[1]!, 'base64')));
    expect(pictureXml(importedQuirks, 0, quirkImages[1]!.id))
      .toContain('descr="/not-read/paired-source.jpeg"');
    const reopenedQuirks = await PptxDocument.open(await importedQuirks.write());
    expect(reopenedQuirks.slides[0]!.shapes.map((shape) => embeddedRasterState(
      reopenedQuirks,
      0,
      shape as ImageModel,
    ))).toEqual(quirkImages.map((image) => embeddedRasterState(importedQuirks, 0, image)));

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const before = packageState(native);
    const shapes = nativeSlide.shapes;
    const invalidCalls: readonly (() => unknown)[] = [
      () => nativeSlide.addImage(
        new Uint8Array([1]),
        { contentType: 'image/png', width: 0 } as never,
      ),
      () => nativeSlide.addImage(PNG_DATA_URI as never, { contentType: 'image/png' }),
      () => nativeSlide.addImage(
        new Uint8Array([1]),
        { contentType: 'image/jpg' } as never,
      ),
      () => nativeSlide.addImage(
        new Uint8Array([1]),
        { contentType: 'image/png', path: '/image.png' } as never,
      ),
      () => nativeSlide.addImage(
        new Uint8Array([1]),
        { contentType: 'image/png', data: PNG_DATA_URI } as never,
      ),
    ];
    for (const invoke of invalidCalls) {
      expect(invoke).toThrow();
      expect(packageState(native)).toEqual(before);
      expect(nativeSlide.shapes).toEqual(shapes);
    }
  }, 20_000);

  it('locks ImageProps source, sizing, and transform divergences against PptxGenJS 4.0.1', async () => {
    const atomIds = [
      ...['data', 'path', 'rotate', 'flipH', 'flipV', 'sizing']
        .map((property) => `interface:ImageProps@property:${property}`),
      ...['type', 'w', 'h', 'x', 'y']
        .map((property) =>
          `inline:interface:ImageProps@property:sizing@property:sizing.${property}`),
      ...['contain', 'cover', 'crop']
        .map((value) =>
          `union:interface:ImageProps@property:sizing@path:sizing.type#${value}`),
    ];
    expect(new Set(atomIds).size).toBe(14);
    for (const property of ['x', 'y', 'w', 'h']) {
      expect(atomIds).not.toContain(`interface:ImageProps@property:${property}`);
    }

    const sizing = new PptxGenJS();
    expect(sizing.version).toBe('4.0.1');
    sizing.defineLayout({ name: 'IMAGE-SIZING-PROBE', width: 10, height: 8 });
    sizing.layout = 'IMAGE-SIZING-PROBE';
    const sizingSlide = sizing.addSlide();
    sizingSlide.addImage({
      data: PNG_DATA_URI,
      objectName: 'Percent contain',
      x: 0,
      y: 0,
      w: 10,
      h: 8,
      sizing: { type: 'contain', w: '30%', h: '40%' },
    });
    sizingSlide.addImage({
      data: PNG_DATA_URI,
      objectName: 'Percent crop',
      x: 0,
      y: 0,
      w: 10,
      h: 8,
      sizing: { type: 'crop', x: '10%', y: '20%', w: '30%', h: '40%' },
    });
    sizingSlide.addImage({
      data: PNG_DATA_URI,
      objectName: 'Missing sizing dimensions',
      x: 0,
      y: 0,
      w: 4,
      h: 3,
      sizing: { type: 'cover' } as never,
    });
    sizingSlide.addImage({
      data: PNG_DATA_URI,
      objectName: 'Outer ratio wins',
      x: 0,
      y: 0,
      w: 16,
      h: 9,
      sizing: { type: 'contain', w: 4, h: 3 },
    });

    const importedSizing = await openPptxGenJSPublicOutput(sizing);
    const imageByName = (name: string) => importedSizing.slides[0]!.shapes.find(
      (shape): shape is ImageModel => shape instanceof ImageModel && shape.name === name,
    )!;
    const contain = imageByName('Percent contain');
    const crop = imageByName('Percent crop');
    const missing = imageByName('Missing sizing dimensions');
    const outerRatio = imageByName('Outer ratio wins');

    expect(contain.transform).toMatchObject({
      x: 0,
      y: 0,
      width: 2_743_200,
      height: 2_926_080,
    });
    expect(pictureXml(importedSizing, 0, contain.id)).toContain(
      '<a:srcRect l="0" r="0" t="-16667" b="-16667"/><a:stretch/>',
    );
    expect(directSourceRectangleRaw(pictureXml(importedSizing, 0, contain.id)))
      .toEqual({ left: 0, top: -16_667, right: 0, bottom: -16_667 });

    expect(crop.transform).toMatchObject({
      x: 0,
      y: 0,
      width: 2_743_200,
      height: 2_926_080,
    });
    expect(pictureXml(importedSizing, 0, crop.id)).toContain(
      '<a:srcRect l="10000" r="60000" t="20000" b="40000"/><a:stretch/>',
    );
    expect(directSourceRectangleRaw(pictureXml(importedSizing, 0, crop.id)))
      .toEqual({ left: 10_000, top: 20_000, right: 60_000, bottom: 40_000 });

    expect(missing.transform).toMatchObject({
      x: 0,
      y: 0,
      width: inches(4),
      height: inches(3),
    });
    expect(pictureXml(importedSizing, 0, missing.id)).toContain(
      '<a:srcRect l="0" r="0" t="0" b="0"/><a:stretch/>',
    );
    expect(directSourceRectangleRaw(pictureXml(importedSizing, 0, outerRatio.id)))
      .toEqual({ left: 0, top: -16_667, right: 0, bottom: -16_667 });

    const intrinsicNative = PptxDocument.create();
    intrinsicNative.addSlide();
    const intrinsic = await intrinsicNative.addImage(0, PNG_DATA_URI, {
      sizing: { type: 'contain', width: inches(4), height: inches(3) },
    });
    expect(directSourceRectangleRaw(pictureXml(intrinsicNative, 0, intrinsic.id)))
      .toEqual({ left: -16_667, top: 0, right: -16_667, bottom: 0 });

    const invalidSizing = new PptxGenJS();
    const invalidSizingSlide = invalidSizing.addSlide();
    expect(() => invalidSizingSlide.addImage({
      data: PNG_DATA_URI,
      sizing: { type: 'bogus', w: 1, h: 1 },
    })).not.toThrow();
    await expect(invalidSizing.write({ outputType: 'nodebuffer', compression: false }))
      .rejects.toThrow(/ImageSizingXml.*is not a function/i);

    const transforms = new PptxGenJS();
    const transformSlide = transforms.addSlide();
    for (const [name, rotate] of [
      ['Rotate 361', 361],
      ['Rotate 721', 721],
      ['Rotate -361', -361],
      ['Rotate string', '45'],
      ['Rotate infinity', Number.POSITIVE_INFINITY],
    ] as const) {
      transformSlide.addImage({
        data: PNG_DATA_URI,
        objectName: name,
        w: 1,
        h: 1,
        rotate,
      });
    }
    transformSlide.addImage({
      data: PNG_DATA_URI,
      objectName: 'Truthy flips',
      w: 1,
      h: 1,
      flipH: 'yes',
      flipV: 1,
    });
    const importedTransforms = await openPptxGenJSPublicOutput(transforms);
    const transformXml = (name: string) => {
      const image = importedTransforms.slides[0]!.shapes.find(
        (shape): shape is ImageModel => shape instanceof ImageModel && shape.name === name,
      )!;
      return pictureXml(importedTransforms, 0, image.id);
    };
    expect(transformXml('Rotate 361')).toContain('<a:xfrm rot="60000">');
    expect(transformXml('Rotate 721')).toContain('<a:xfrm rot="21660000">');
    expect(transformXml('Rotate -361')).toContain('<a:xfrm rot="-21660000">');
    expect(transformXml('Rotate string')).toContain('<a:xfrm rot="2700000">');
    expect(transformXml('Rotate infinity')).toContain('<a:xfrm rot="Infinity">');
    expect(transformXml('Truthy flips')).toContain('<a:xfrm flipH="1" flipV="1">');

    const invalidSources = new PptxGenJS();
    const invalidSourceSlide = invalidSources.addSlide();
    invalidSourceSlide.addImage({
      data: 'data:image/png;base64,',
      objectName: 'Empty source',
    });
    invalidSourceSlide.addImage({
      data: 'data:image/png;base64,%%%%',
      objectName: 'Junk source',
    });
    const importedInvalidSources = await openPptxGenJSPublicOutput(invalidSources);
    const invalidImages = importedInvalidSources.slides[0]!.shapes as readonly ImageModel[];
    expect(invalidImages.map((image) => importedInvalidSources.opcPackage
      .requirePart(image.sourcePartUri!).bytes.byteLength)).toEqual([0, 0]);

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const before = packageState(native);
    const beforeShapes = nativeSlide.shapes;
    const rasterBytes = new Uint8Array(
      Buffer.from(PNG_DATA_URI.split(',')[1]!, 'base64'),
    );
    for (const options of [
      { contentType: 'image/png', rotation: degrees(361) },
      { contentType: 'image/png', rotation: Number.POSITIVE_INFINITY },
      { contentType: 'image/png', flipHorizontal: 'yes' },
      { contentType: 'image/png', flipVertical: 1 },
    ]) {
      expect(() => nativeSlide.addImage(rasterBytes, options as never)).toThrow();
      expect(packageState(native)).toEqual(before);
      expect(nativeSlide.shapes).toEqual(beforeShapes);
    }
    for (const source of ['data:image/png;base64,', 'data:image/png;base64,%%%%']) {
      await expect(native.addImage(0, source)).rejects.toThrow();
      expect(packageState(native)).toEqual(before);
      expect(nativeSlide.shapes).toEqual(beforeShapes);
    }
    for (const value of [
      { type: 'bogus', width: inches(1), height: inches(1) },
      { type: 'cover' },
    ]) {
      await expect(native.addImage(0, PNG_DATA_URI, { sizing: value } as never))
        .rejects.toThrow();
      expect(packageState(native)).toEqual(before);
      expect(nativeSlide.shapes).toEqual(beforeShapes);
    }
  }, 20_000);

  it('closes PptxGenJS core content and primitive inputs through strict native state', async () => {
    const classifications = [
      ...[
        'interface:TextProps@property:text',
        'union:interface:TableCell@property:text#string',
        'union:method:Slide#addText@path:text#string',
        'union:Coord#${number}%',
        'union:Margin#number',
      ].map((id) => ({ id, status: 'supported' })),
      ...[
        'interface:TableCell@property:options',
        'interface:TableCell@property:text',
        'interface:TextProps@property:options',
        'union:interface:TableCell@property:text#TableCell[]',
        'union:method:Slide#addText@path:text#TextProps[]',
        'union:Color#HexColor',
        'union:Color#ThemeColor',
        'union:Coord#number',
        'union:Margin#[number,number,number,number]',
      ].map((id) => ({ id, status: 'deliberate-difference' })),
    ].sort((left, right) => left.id.localeCompare(right.id));
    expect(classifications).toHaveLength(14);
    expect(new Set(classifications.map(({ id }) => id)).size).toBe(14);
    expect({
      supported: classifications.filter(({ status }) => status === 'supported').length,
      deliberate: classifications.filter(
        ({ status }) => status === 'deliberate-difference',
      ).length,
    }).toEqual({ supported: 5, deliberate: 9 });

    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.defineLayout({ name: 'CORE_PRIMITIVES_10X8', width: 10, height: 8 });
    generated.layout = 'CORE_PRIMITIVES_10X8';
    const first = generated.addSlide();
    first.addText('Plain string', {
      objectName: 'Core plain string',
      x: 1,
      y: '10%',
      w: '30%',
      h: 1,
      margin: 10,
    });
    first.addText([
      { text: 'Hex run', options: { bold: true, color: '112233' } },
      {
        text: ' Theme run',
        options: { italic: true, color: generated.SchemeColor.accent2 },
      },
    ], {
      objectName: 'Core rich runs',
      x: '10%',
      y: 2,
      w: 5,
      h: '20%',
      margin: [1, 2, 3, 4],
    });
    first.addText('Zero margin', {
      objectName: 'Core zero margin',
      x: 1,
      y: 4,
      w: 2,
      h: 0.5,
      margin: 0,
    });
    const second = generated.addSlide();
    second.addTable([[
      'Bare string cell',
      {
        text: 'Structured plain cell',
        options: { bold: true, color: '445566', margin: 0 },
      },
      {
        text: [
          { text: 'Rich cell hex', options: { bold: true, color: 'AABBCC' } },
          {
            text: ' rich cell theme',
            options: { italic: true, color: generated.SchemeColor.accent3 },
          },
        ],
        options: { margin: [1, 2, 3, 4] },
      },
    ]], {
      objectName: 'Core primitive table',
      x: 1,
      y: 1,
      w: 8,
      h: 1.5,
      margin: 0,
    });

    const imported = await openPptxGenJSPublicOutput(generated);
    const [plain, rich, zero] = imported.slides[0]!.shapes as readonly ShapeModel[];
    const table = imported.slides[1]!.shapes[0] as TableModel;
    expect([plain!.text, rich!.text, zero!.text]).toEqual([
      'Plain string',
      'Hex run Theme run',
      'Zero margin',
    ]);
    expect(plain!.transform).toEqual({
      x: 914_400,
      y: 731_520,
      width: 2_743_200,
      height: 914_400,
      rotation: 0,
      flipHorizontal: false,
      flipVertical: false,
    });
    expect(rich!.transform).toEqual({
      x: 914_400,
      y: 1_828_800,
      width: 4_572_000,
      height: 1_463_040,
      rotation: 0,
      flipHorizontal: false,
      flipVertical: false,
    });
    expect(plain!.textMargins).toEqual({ top: 10, right: 10, bottom: 10, left: 10 });
    expect(rich!.textMargins).toEqual({ top: 4, right: 2, bottom: 3, left: 1 });
    expect(zero!.textMargins).toEqual({ top: 0, right: 0, bottom: 0, left: 0 });
    expect(rich!.richText[0]!.runs.map(({ text, style }) => ({
      text,
      color: style?.color,
    }))).toEqual([
      { text: 'Hex run', color: { kind: 'srgb', value: '112233' } },
      { text: ' Theme run', color: { kind: 'scheme', value: 'accent2' } },
    ]);
    expect(table).toBeInstanceOf(TableModel);
    expect(table.rows[0]!.cells.map(({ text }) => text)).toEqual([
      'Bare string cell',
      'Structured plain cell',
      'Rich cell hex rich cell theme',
    ]);
    expect(table.rows[0]!.cells[1]!.richText[0]!.runs[0]!.style?.color).toEqual({
      kind: 'srgb',
      value: '445566',
    });
    expect(table.rows[0]!.cells[2]!.richText[0]!.runs.map(({ text, style }) => ({
      text,
      color: style?.color,
    }))).toEqual([
      { text: 'Rich cell hex', color: { kind: 'srgb', value: 'AABBCC' } },
      { text: ' rich cell theme', color: { kind: 'scheme', value: 'accent3' } },
    ]);

    const ownerXml = (xml: string, name: string, tag = 'p:sp'): string => {
      const identityOffset = xml.indexOf(`name="${name}"`);
      const start = xml.lastIndexOf(`<${tag}`, identityOffset);
      const end = xml.indexOf(`</${tag}>`, identityOffset);
      expect(identityOffset).toBeGreaterThanOrEqual(0);
      expect(start).toBeGreaterThanOrEqual(0);
      expect(end).toBeGreaterThan(identityOffset);
      return xml.slice(start, end + tag.length + 3);
    };
    const firstXml = slideXml(imported, 0);
    const secondXml = slideXml(imported, 1);
    const publicRichXml = ownerXml(firstXml, 'Core rich runs');
    expect(ownerXml(firstXml, 'Core plain string')).toContain(
      'lIns="127000" tIns="127000" rIns="127000" bIns="127000"',
    );
    expect(publicRichXml).toContain(
      'lIns="12700" tIns="50800" rIns="25400" bIns="38100"',
    );
    expect(ownerXml(firstXml, 'Core zero margin')).toContain(
      'lIns="0" tIns="0" rIns="0" bIns="0"',
    );
    expect(publicRichXml).toContain('<a:srgbClr val="112233"/>');
    expect(publicRichXml).toContain('<a:schemeClr val="accent2"/>');
    expect(ownerXml(secondXml, 'Core primitive table', 'p:graphicFrame')).toContain(
      '<a:schemeClr val="accent3"/>',
    );

    const native = PptxDocument.create({
      slideSize: { width: inches(10), height: inches(8) },
    });
    const nativeSlide = native.addSlide();
    const nativePercent = nativeSlide.addText('Native explicit units', {
      x: inches(1),
      y: '10%',
      width: '30%',
      height: inches(1),
    });
    nativeSlide.addRichText([{
      runs: [
        { text: 'Native hex', style: { color: { kind: 'srgb', value: '112233' } } },
        {
          text: ' native theme',
          style: { color: { kind: 'scheme', value: 'accent2' } },
        },
      ],
    }], {
      name: 'Native documented tuple',
      margin: [1, 2, 3, 4],
    });
    expect(nativePercent.transform).toEqual(plain!.transform);
    expect(ownerXml(slideXml(native, 0), 'Native documented tuple')).toContain(
      'lIns="50800" tIns="12700" rIns="25400" bIns="38100"',
    );

    plain!.text = 'Plain edited';
    rich!.richText = [{
      runs: [
        {
          text: 'Edited hex',
          style: { bold: true, color: { kind: 'srgb', value: 'ABCDEF' } },
        },
        {
          text: ' Edited theme',
          style: { italic: true, color: { kind: 'scheme', value: 'accent4' } },
        },
      ],
    }];
    table.setCellRichText(0, 2, [{
      runs: [
        { text: 'Edited table', style: { color: { kind: 'srgb', value: '123456' } } },
        { text: ' theme', style: { color: { kind: 'scheme', value: 'accent5' } } },
      ],
    }]);
    const reopened = await PptxDocument.open(await imported.write({
      compatibility: 'powerpoint-2010',
    }));
    const reopenedTexts = reopened.slides[0]!.shapes as readonly ShapeModel[];
    expect(reopenedTexts.map((shape) => shape.text)).toEqual([
      'Plain edited',
      'Edited hex Edited theme',
      'Zero margin',
    ]);
    expect((reopened.slides[1]!.shapes[0] as TableModel)
      .rows[0]!.cells[2]!.richText[0]!.runs.map(({ text, style }) => ({
        text,
        color: style?.color,
      }))).toEqual([
      { text: 'Edited table', color: { kind: 'srgb', value: '123456' } },
      { text: ' theme', color: { kind: 'scheme', value: 'accent5' } },
    ]);
    expect(imported.diagnostics).toEqual([]);
    expect(reopened.diagnostics).toEqual([]);
  }, 20_000);

  it('closes PptxGenJS hyperlink owners through shared strict native state', async () => {
    const classifications = [
      ...[
        'interface:HyperlinkProps@property:slide',
        'interface:HyperlinkProps@property:tooltip',
        'interface:HyperlinkProps@property:url',
        'interface:ImageProps@property:hyperlink',
        'interface:ShapeProps@property:hyperlink',
      ].map((id) => ({ id, status: 'supported' })),
      {
        id: 'interface:TextPropsOptions@property:hyperlink',
        status: 'deliberate-difference',
      },
    ].sort((left, right) => left.id.localeCompare(right.id));
    expect(classifications).toHaveLength(6);
    expect(new Set(classifications.map(({ id }) => id)).size).toBe(6);
    expect({
      supported: classifications.filter(({ status }) => status === 'supported').length,
      deliberate: classifications.filter(
        ({ status }) => status === 'deliberate-difference',
      ).length,
    }).toEqual({ supported: 5, deliberate: 1 });

    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const first = generated.addSlide();
    generated.addSlide();
    generated.addSlide();
    const shapeLink = {
      url: 'https://example.com/shape?a=1&b=2',
      tooltip: 'Shape & tip',
    };
    const imageLink = { slide: 3, tooltip: '' };
    const textLink = { url: 'mailto:text@example.com', tooltip: 'Text tip' };
    const richOuterLink = { url: 'https://example.com/rich-outer' };
    first.addShape(generated.ShapeType.rect!, {
      objectName: 'Hyperlink Shape',
      x: 0.5,
      y: 0.5,
      w: 2,
      h: 0.5,
      hyperlink: shapeLink,
    });
    first.addImage({
      objectName: 'Hyperlink Image',
      data: PNG_DATA_URI,
      x: 0.5,
      y: 1.25,
      w: 2,
      h: 1,
      hyperlink: imageLink,
    });
    first.addText('Plain hyperlink', {
      objectName: 'Hyperlink Text',
      x: 0.5,
      y: 2.5,
      w: 2,
      h: 0.5,
      hyperlink: textLink,
    });
    first.addText([
      { text: 'Rich outer one', options: { bold: true } },
      { text: ' two', options: { italic: true } },
    ], {
      objectName: 'Hyperlink Rich Outer',
      x: 0.5,
      y: 3.25,
      w: 2,
      h: 0.5,
      hyperlink: richOuterLink,
    });

    const imported = await openPptxGenJSPublicOutput(generated);
    const [shape, image, text, richOuter] = imported.slides[0]!.shapes;
    expect((shape as ShapeModel).hyperlink).toEqual({
      url: 'https://example.com/shape?a=1&b=2',
      tooltip: 'Shape & tip',
    });
    expect((image as ImageModel).hyperlink).toEqual({ slide: 3, tooltip: '' });
    expect((text as ShapeModel).hyperlink).toEqual({
      url: 'mailto:text@example.com',
      tooltip: 'Text tip',
    });
    expect((richOuter as ShapeModel).hyperlink).toBeUndefined();
    expect(Object.hasOwn(shapeLink, '_rId')).toBe(true);
    expect(Object.hasOwn(imageLink, '_rId')).toBe(true);
    expect(Object.hasOwn(textLink, '_rId')).toBe(true);
    const importedXml = slideXml(imported, 0);
    expect(importedXml.match(/r:id="rIdundefined"/g)).toHaveLength(3);
    expect(pictureXml(imported, 0, (image as ImageModel).id)).toContain(
      'tooltip="" action="ppaction://hlinksldjump"',
    );

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    native.addSlide();
    native.addSlide();
    const nativeInput: { slide: number; tooltip?: string } = { slide: 3, tooltip: '' };
    const nativeImage = nativeSlide.addImage(
      new Uint8Array(Buffer.from(PNG_DATA_URI.split(',')[1]!, 'base64')),
      { contentType: 'image/png', hyperlink: nativeInput },
    );
    nativeInput.slide = 1;
    delete nativeInput.tooltip;
    expect(nativeImage.hyperlink).toEqual({ slide: 3, tooltip: '' });
    expect(Object.hasOwn(nativeInput, '_rId')).toBe(false);
    expect(slideXml(native, 0)).not.toContain('rIdundefined');
    const reopened = await PptxDocument.open(await native.write({
      compatibility: 'powerpoint-2010',
    }));
    expect((reopened.slides[0]!.shapes[0] as ImageModel).hyperlink)
      .toEqual({ slide: 3, tooltip: '' });
    expect(native.diagnostics).toEqual([]);
    expect(reopened.diagnostics).toEqual([]);
  }, 20_000);

  it('closes inherited data and path declarations against real source owners', async () => {
    const classifications = [
      ...['data', 'path'].map((property) => ({
        id: `interface:DataOrPathProps@property:${property}`,
        status: 'deliberate-difference',
      })),
      ...['data', 'path'].map((property) => ({
        id: `interface:TextPropsOptions@property:${property}`,
        status: 'defect-excluded',
      })),
    ].sort((left, right) => left.id.localeCompare(right.id));
    expect(classifications).toHaveLength(4);
    expect(new Set(classifications.map(({ id }) => id)).size).toBe(4);
    expect(classifications.map(({ status }) => status).sort()).toEqual([
      'defect-excluded',
      'defect-excluded',
      'deliberate-difference',
      'deliberate-difference',
    ]);

    const directory = await mkdtemp(join(tmpdir(), 'pptxgenjs-data-path-'));
    const imagePath = join(directory, 'path.gif');
    const audioPath = join(directory, 'path.mp3');
    const imageDataBytes = new Uint8Array(Buffer.from(PNG_DATA_URI.split(',')[1]!, 'base64'));
    const imagePathBytes = new Uint8Array(Buffer.from(GIF_DATA_URI.split(',')[1]!, 'base64'));
    const audioData = 'data:audio/mpeg;base64,AQIDBA==';
    const audioDataBytes = Uint8Array.of(1, 2, 3, 4);
    const audioPathBytes = Uint8Array.of(9, 10, 11, 12, 13);
    const dataSentinel = 'TEXT_DATA_SENTINEL_SHOULD_NOT_SERIALIZE';
    const pathSentinel = 'TEXT_PATH_SENTINEL_SHOULD_NOT_SERIALIZE';
    try {
      await writeFile(imagePath, imagePathBytes);
      await writeFile(audioPath, audioPathBytes);
      const generated = new PptxGenJS();
      generated.addSlide().addImage({
        data: PNG_DATA_URI, path: imagePath, objectName: 'image-both', x: 1, y: 1, w: 2, h: 2,
      });
      generated.addSlide().addImage({
        path: imagePath, objectName: 'image-path', x: 1, y: 1, w: 2, h: 2,
      });
      generated.addSlide().addImage({
        data: PNG_DATA_URI, objectName: 'image-data', x: 1, y: 1, w: 2, h: 2,
      });
      generated.addSlide().addMedia({
        type: 'audio', data: audioData, path: audioPath, cover: PNG_DATA_URI, extn: 'mp3',
        objectName: 'media-both', x: 1, y: 1, w: 2, h: 1,
      });
      generated.addSlide().addMedia({
        type: 'audio', path: audioPath, cover: PNG_DATA_URI,
        objectName: 'media-path', x: 1, y: 1, w: 2, h: 1,
      });
      generated.addSlide().addMedia({
        type: 'audio', data: audioData, cover: PNG_DATA_URI, extn: 'mp3',
        objectName: 'media-data', x: 1, y: 1, w: 2, h: 1,
      });
      generated.addSlide().addText('Plain inert source', {
        objectName: 'plain-byte-control', x: 1, y: 1, w: 4, h: 1,
      });
      generated.addSlide().addText('Plain inert source', {
        data: dataSentinel, path: pathSentinel,
        objectName: 'plain-byte-control', x: 1, y: 1, w: 4, h: 1,
      });
      const richRuns = [
        { text: 'Rich inert ', options: { bold: true } },
        { text: 'source', options: { color: '112233' } },
      ];
      generated.addSlide().addText(richRuns, {
        objectName: 'rich-byte-control', x: 1, y: 1, w: 4, h: 1,
      });
      generated.addSlide().addText(richRuns, {
        data: dataSentinel, path: pathSentinel,
        objectName: 'rich-byte-control', x: 1, y: 1, w: 4, h: 1,
      });

      const imported = await openPptxGenJSPublicOutput(generated);
      const importedImages = imported.slides.slice(0, 3)
        .map((slide) => slide.shapes[0] as ImageModel);
      expect(importedImages.map((image, index) =>
        embeddedRasterState(imported, index, image).bytes)).toEqual([
        [...imageDataBytes],
        [...imagePathBytes],
        [...imageDataBytes],
      ]);
      expect(embeddedRasterState(imported, 0, importedImages[0]!).altText).toBe(imagePath);

      const importedMedia = imported.slides.slice(3, 6)
        .map((_, index) => embeddedMediaStates(imported, index + 3)[0]!);
      expect(importedMedia.map(({ mediaBytes }) => mediaBytes)).toEqual([
        [...audioDataBytes],
        [...audioPathBytes],
        [...audioDataBytes],
      ]);
      expect(slideXml(imported, 3)).not.toContain(audioPath);

      const normalizeTextXml = (xml: string) => xml
        .replace(/<p:cSld name="Slide \d+"/gu, '<p:cSld name="Slide #"')
        .replace(/<p:cNvPr id="\d+"/gu, '<p:cNvPr id="#"');
      expect(normalizeTextXml(slideXml(imported, 6)))
        .toBe(normalizeTextXml(slideXml(imported, 7)));
      expect(normalizeTextXml(slideXml(imported, 8)))
        .toBe(normalizeTextXml(slideXml(imported, 9)));
      for (const part of imported.opcPackage.parts) {
        const text = new TextDecoder().decode(part.bytes);
        expect(text).not.toContain(dataSentinel);
        expect(text).not.toContain(pathSentinel);
      }

      const native = PptxDocument.create();
      native.addSlide();
      native.addSlide();
      native.addSlide();
      const nativeImage = await native.addImage(0, imageDataBytes, { contentType: 'image/png' });
      const nativeAudio = await native.addAudio(1, audioData, { poster: PNG_DATA_URI });
      const nativeText = native.slides[2]!.addText('Plain inert source');
      expect(embeddedRasterState(native, 0, nativeImage).bytes).toEqual([...imageDataBytes]);
      expect(embeddedMediaStates(native, 1)[0]!.mediaBytes).toEqual([...audioDataBytes]);
      for (const owner of [nativeImage, nativeAudio, nativeText]) {
        expect('data' in owner).toBe(false);
        expect('path' in owner).toBe(false);
      }
      const reopened = await PptxDocument.open(await native.write({
        compatibility: 'powerpoint-2010',
      }));
      expect(embeddedRasterState(reopened, 0, reopened.slides[0]!.shapes[0] as ImageModel).bytes)
        .toEqual([...imageDataBytes]);
      expect(embeddedMediaStates(reopened, 1)[0]!.mediaBytes).toEqual([...audioDataBytes]);
      expect(native.diagnostics).toEqual([]);
      expect(reopened.diagnostics).toEqual([]);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  }, 20_000);

  it('closes placeholder text style fields through canonical text owners', async () => {
    const classifications = [
      { id: 'interface:PlaceholderProps@property:align', status: 'supported' },
      { id: 'interface:PlaceholderProps@property:margin', status: 'deliberate-difference' },
      { id: 'interface:PlaceholderProps@property:transparency', status: 'supported' },
      { id: 'interface:PlaceholderProps@property:valign', status: 'supported' },
    ].sort((left, right) => left.id.localeCompare(right.id));
    expect(classifications).toHaveLength(4);
    expect(new Set(classifications.map(({ id }) => id)).size).toBe(4);
    expect(classifications.map(({ status }) => status).sort()).toEqual([
      'deliberate-difference',
      'supported',
      'supported',
      'supported',
    ]);

    const alignMap = { left: 'l', center: 'ctr', right: 'r', justify: 'just' } as const;
    const valignMap = { top: 't', middle: 'ctr', bottom: 'b' } as const;
    const cases: readonly {
      readonly family: 'align' | 'margin' | 'transparency' | 'valign';
      readonly label: string;
      readonly value?: unknown;
      readonly omitted?: true;
    }[] = [
      ...Object.keys(alignMap).map((value) => ({ family: 'align' as const, label: value, value })),
      ...Object.keys(valignMap).map((value) => ({ family: 'valign' as const, label: value, value })),
      { family: 'margin', label: 'scalar-zero', value: 0 },
      { family: 'margin', label: 'scalar-eight', value: 8 },
      { family: 'margin', label: 'tuple-asymmetric', value: [1, 2, 3, 4] },
      { family: 'transparency', label: 'omitted', omitted: true },
      { family: 'transparency', label: 'zero', value: 0 },
      { family: 'transparency', label: 'fractional', value: 12.5 },
      { family: 'transparency', label: 'hundred', value: 100 },
    ];
    expect(cases).toHaveLength(14);

    const decodeEntities = (value: string): string => value
      .replaceAll('&amp;', '&')
      .replaceAll('&quot;', '"')
      .replaceAll('&apos;', "'")
      .replaceAll('&lt;', '<')
      .replaceAll('&gt;', '>');
    const attributes = (value: string): Record<string, string> => Object.fromEntries(
      [...value.matchAll(/([\w:]+)="([^"]*)"/g)]
        .map(([, key, entry]) => [key!, decodeEntities(entry!)]),
    );
    const ownerMap = (xml: string): Map<string, string> => new Map(
      [...xml.matchAll(/<p:sp(?:\s[^>]*)?>[\s\S]*?<\/p:sp>/g)].map(([owner]) => {
        const name = attributes(owner!.match(/<p:cNvPr\b([^>]*)>/)?.[1] ?? '').name;
        return [name ?? '', owner!];
      }),
    );
    const ownerState = (owner: string) => {
      const body = attributes(owner.match(/<a:bodyPr\b([^>]*)>/)?.[1] ?? '');
      const paragraph = attributes(owner.match(/<a:pPr\b([^>]*)>/)?.[1] ?? '');
      return {
        align: paragraph.algn,
        anchor: body.anchor,
        margins: {
          left: body.lIns === undefined ? undefined : Number(body.lIns),
          top: body.tIns === undefined ? undefined : Number(body.tIns),
          right: body.rIns === undefined ? undefined : Number(body.rIns),
          bottom: body.bIns === undefined ? undefined : Number(body.bIns),
        },
        alphas: [...new Set([...owner.matchAll(/<a:alpha\s+val="([^"]+)"\s*\/>/g)]
          .map(([, value]) => Number(value)))],
      };
    };
    const caseId = ({ family, label }: (typeof cases)[number]): string =>
      `${family}-${label}`;
    const objectName = (entry: (typeof cases)[number]): string =>
      `Placeholder style ${caseId(entry)}`;

    const publicObjects: PptxGenJSMasterObject[] = cases.map((entry, index) => {
      const options: PptxGenJSPlaceholderOptions = {
        name: caseId(entry),
        objectName: objectName(entry),
        type: 'body',
        x: 0.25 + (index % 4) * 3,
        y: 0.25 + Math.floor(index / 4) * 1.6,
        w: 2.75,
        h: 1.2,
        color: '336699',
      };
      if (!entry.omitted) options[entry.family] = entry.value;
      return { placeholder: { text: `Prompt ${caseId(entry)}`, options } };
    });
    const masterInput = { title: 'PLACEHOLDER-TEXT-STYLE-4', objects: publicObjects };
    const masterBefore = JSON.stringify(masterInput);
    const source = new PptxGenJS();
    expect(source.version).toBe('4.0.1');
    source.layout = 'LAYOUT_WIDE';
    source.defineSlideMaster(masterInput);
    expect(JSON.stringify(masterInput)).toBe(masterBefore);
    const publicSlide = source.addSlide({ masterName: masterInput.title });
    const populationInputs = cases.map((entry) => ({ placeholder: caseId(entry) }));
    cases.forEach((entry, index) => {
      publicSlide.addText(`Populated ${caseId(entry)}`, populationInputs[index]!);
    });
    expect(populationInputs.every((options) =>
      JSON.stringify(options) === JSON.stringify({
        placeholder: options.placeholder,
        bullet: false,
      }))).toBe(true);

    const imported = await openPptxGenJSPublicOutput(source);
    const importedLayout = imported.layouts.find(({ name }) => name === masterInput.title)!;
    const importedLayoutXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(importedLayout.partUri).bytes,
    );
    const publicLayoutOwners = ownerMap(importedLayoutXml);
    const publicSlideOwners = ownerMap(slideXml(imported, 0));
    const publicStates = new Map(cases.map((entry) => [caseId(entry), {
      layout: ownerState(publicLayoutOwners.get(objectName(entry)) ?? ''),
      populated: ownerState(publicSlideOwners.get(objectName(entry)) ?? ''),
    }]));
    expect(publicStates.size).toBe(14);
    expect(cases.every((entry) => publicLayoutOwners.has(objectName(entry)))).toBe(true);
    expect(cases.every((entry) => publicSlideOwners.has(objectName(entry)))).toBe(true);

    for (const [token, expected] of Object.entries(alignMap)) {
      const state = publicStates.get(`align-${token}`)!;
      expect(state.layout.align, token).toBe(expected);
      expect(state.populated.align, token).toBe(expected);
    }
    for (const [token, expected] of Object.entries(valignMap)) {
      const state = publicStates.get(`valign-${token}`)!;
      expect(state.layout.anchor, token).toBeUndefined();
      expect(state.populated.anchor, token).toBe(expected);
    }
    const zeroMargin = publicStates.get('margin-scalar-zero')!;
    const positiveMargin = publicStates.get('margin-scalar-eight')!;
    const tupleMargin = publicStates.get('margin-tuple-asymmetric')!;
    expect([zeroMargin, positiveMargin, tupleMargin].every(({ layout }) =>
      Object.values(layout.margins).every((value) => value === undefined))).toBe(true);
    expect(zeroMargin.populated.margins).toEqual({ left: 0, top: 0, right: 0, bottom: 0 });
    expect(positiveMargin.populated.margins).toEqual({
      left: 101_600,
      top: 101_600,
      right: 101_600,
      bottom: 101_600,
    });
    expect(tupleMargin.populated.margins).toEqual({
      left: 12_700,
      top: 50_800,
      right: 25_400,
      bottom: 38_100,
    });
    expect(publicStates.get('transparency-omitted')).toMatchObject({
      layout: { alphas: [] }, populated: { alphas: [] },
    });
    expect(publicStates.get('transparency-zero')).toMatchObject({
      layout: { alphas: [] }, populated: { alphas: [] },
    });
    expect(publicStates.get('transparency-fractional')).toMatchObject({
      layout: { alphas: [87_500] }, populated: { alphas: [87_500] },
    });
    expect(publicStates.get('transparency-hundred')).toMatchObject({
      layout: { alphas: [0] }, populated: { alphas: [0] },
    });

    const packageSnapshot = (document: PptxDocument): string => JSON.stringify({
      parts: document.opcPackage.parts.map(
        ({ uri, contentType, bytes, relationships }) => ({
          uri,
          contentType,
          bytes: [...bytes],
          relationships,
        }),
      ),
      mutations: [...document.opcPackage.mutations],
    });
    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeLayout = native.layouts[0]!;
    for (const [index, entry] of cases.entries()) {
      const nativeOptions: Record<string, unknown> & {
        name: string;
        type: 'body';
        index: number;
      } = {
        name: objectName(entry),
        type: 'body',
        index: 300 + index,
      };
      let value: string | readonly {
        readonly runs: readonly {
          readonly text: string;
          readonly style?: { readonly color?: unknown; readonly transparency?: number };
        }[];
      }[] = `Prompt ${caseId(entry)}`;
      if (entry.family === 'transparency') {
        value = [{
          runs: [{
            text: `Prompt ${caseId(entry)}`,
            style: {
              color: { kind: 'srgb', value: '336699' },
              ...(!entry.omitted ? { transparency: entry.value as number } : {}),
            },
          }],
        }];
      } else if (entry.family === 'margin') {
        nativeOptions.margin = entry.label === 'tuple-asymmetric'
          ? [4, 2, 3, 1]
          : entry.value;
      } else {
        nativeOptions[entry.family] = entry.value;
      }
      nativeLayout.addPlaceholder(value as never, nativeOptions as never);
    }

    const detachedValue = [{
      runs: [{ text: 'Detached', style: { transparency: 25 } }],
    }];
    const detachedOptions: {
      name: string;
      type: 'body';
      index: number;
      align: 'center' | 'right';
      margin: [number, number, number, number];
      valign: 'bottom' | 'top';
    } = {
      name: 'Detached placeholder text style',
      type: 'body',
      index: 399,
      align: 'center',
      margin: [1, 2, 3, 4],
      valign: 'bottom',
    };
    const detached = nativeLayout.addPlaceholder(detachedValue, detachedOptions);
    const detachedSnapshot = packageSnapshot(native);
    detachedValue[0]!.runs[0]!.style.transparency = 75;
    detachedOptions.align = 'right';
    detachedOptions.margin[0] = 9;
    detachedOptions.valign = 'top';
    expect(packageSnapshot(native)).toBe(detachedSnapshot);
    expect(detached.richText[0]!.runs[0]!.style?.transparency).toBe(25);
    expect(detached.richText[0]!.align).toBe('center');
    expect(detached.textMargins).toEqual({ top: 1, right: 2, bottom: 3, left: 4 });
    expect(detached.verticalAlignment).toBe('bottom');

    const beforeInvalid = packageSnapshot(native);
    for (const invoke of [
      () => nativeLayout.addPlaceholder('Invalid align', {
        name: 'Invalid align', type: 'body', align: 'centered' as never,
      }),
      () => nativeLayout.addPlaceholder('Invalid margin', {
        name: 'Invalid margin', type: 'body', margin: [1, 2, 3] as never,
      }),
      () => nativeLayout.addPlaceholder('Invalid valign', {
        name: 'Invalid valign', type: 'body', valign: 'center' as never,
      }),
      () => nativeLayout.addPlaceholder([{
        runs: [{ text: 'Invalid transparency', style: { transparency: 101 } }],
      }], { name: 'Invalid transparency', type: 'body' }),
    ]) expect(invoke).toThrow();
    expect(packageSnapshot(native)).toBe(beforeInvalid);

    const nativeSlide = native.addSlide({ masterName: nativeLayout.name });
    for (const entry of cases) {
      const populationOptions: Record<string, unknown> = { placeholder: objectName(entry) };
      if (entry.family === 'transparency') {
        nativeSlide.addRichText([{
          runs: [{
            text: `Populated ${caseId(entry)}`,
            style: {
              color: { kind: 'srgb', value: '336699' },
              ...(!entry.omitted ? { transparency: entry.value as number } : {}),
            },
          }],
        }], populationOptions);
      } else {
        populationOptions[entry.family] = entry.family === 'margin'
            && entry.label === 'tuple-asymmetric'
          ? [4, 2, 3, 1]
          : entry.value;
        nativeSlide.addText(`Populated ${caseId(entry)}`, populationOptions);
      }
    }

    const output = await native.write({ compatibility: 'powerpoint-2010' });
    const reopened = await PptxDocument.open(output);
    const nativeOwners = ownerMap(slideXml(reopened, 0));
    const nativeStates = new Map(cases.map((entry) => [
      caseId(entry),
      ownerState(nativeOwners.get(objectName(entry)) ?? ''),
    ]));
    for (const [token, expected] of Object.entries(alignMap)) {
      expect(nativeStates.get(`align-${token}`)?.align, token).toBe(expected);
    }
    for (const [token, expected] of Object.entries(valignMap)) {
      expect(nativeStates.get(`valign-${token}`)?.anchor, token).toBe(expected);
    }
    expect(nativeStates.get('margin-scalar-zero')?.margins)
      .toEqual(zeroMargin.populated.margins);
    expect(nativeStates.get('margin-scalar-eight')?.margins)
      .toEqual(positiveMargin.populated.margins);
    expect(nativeStates.get('margin-tuple-asymmetric')?.margins)
      .toEqual(tupleMargin.populated.margins);
    expect(nativeStates.get('transparency-omitted')?.alphas).toEqual([]);
    expect(nativeStates.get('transparency-zero')?.alphas).toEqual([100_000]);
    expect(nativeStates.get('transparency-fractional')?.alphas).toEqual([87_500]);
    expect(nativeStates.get('transparency-hundred')?.alphas).toEqual([0]);
    expect(native.diagnostics).toEqual([]);
    expect(reopened.diagnostics).toEqual([]);
  }, 30_000);

  it('closes residual PptxGenJS shape declarations against strict direct adjustment state', async () => {
    const classifications = [
      { id: 'interface:ShapeProps@property:align', status: 'defect-excluded' },
      { id: 'interface:ShapeProps@property:angleRange', status: 'deliberate-difference' },
      { id: 'interface:ShapeProps@property:arcThicknessRatio', status: 'deliberate-difference' },
    ].sort((left, right) => left.id.localeCompare(right.id));
    expect(classifications).toHaveLength(3);
    expect(new Set(classifications.map(({ id }) => id)).size).toBe(3);
    expect(classifications.map(({ status }) => status).sort()).toEqual([
      'defect-excluded',
      'deliberate-difference',
      'deliberate-difference',
    ]);

    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const alignTokens = [
      undefined,
      'left',
      'center',
      'right',
      'justify',
      'runtime-invalid-align',
    ] as const;
    const alignInputs: (PptxGenJSShapeOptions & {
      align?: string;
      line?: { type: string };
    })[] = alignTokens.map((align) => ({
      objectName: 'Inert shape alignment',
      x: 1,
      y: 1,
      w: 3,
      h: 2,
      ...(align === undefined ? {} : { align }),
    }));
    for (const options of alignInputs) {
      generated.addSlide().addShape(generated.ShapeType.rect!, options);
    }
    const geometrySlide = generated.addSlide();
    const fixtures = [
      {
        name: 'Residual pie angles',
        type: 'pie',
        options: { angleRange: [270, 0] },
        expected: [
          { name: 'adj1', value: 16_200_000 },
          { name: 'adj2', value: 0 },
        ],
      },
      {
        name: 'Residual block arc thickness',
        type: 'blockArc',
        options: { angleRange: [270, 0], arcThicknessRatio: 0.5 },
        expected: [
          { name: 'adj1', value: 16_200_000 },
          { name: 'adj2', value: 0 },
          { name: 'adj3', value: 25_000 },
        ],
      },
    ] as const;
    for (const fixture of fixtures) {
      geometrySlide.addShape(generated.ShapeType[fixture.type]!, {
        objectName: fixture.name,
        ...fixture.options,
      });
    }
    expect(alignInputs.map(({ align, line }) => ({ align, line }))).toEqual([
      { align: undefined, line: { type: 'none' } },
      ...alignTokens.slice(1).map((align) => ({ align, line: { type: 'none' } })),
    ]);

    const imported = await openPptxGenJSPublicOutput(generated);
    const alignOwnerXml = imported.slides.slice(0, alignTokens.length).map(
      (slide, index) => shapeXml(imported, index, slide.shapes[0]!.id),
    );
    expect(new Set(alignOwnerXml).size).toBe(1);
    expect(alignOwnerXml[0]).not.toContain('<p:txBody>');
    expect(alignOwnerXml[0]).not.toMatch(/\balgn="/);
    const alignRelationshipKinds = imported.slides.slice(0, alignTokens.length)
      .map((slide) => slide.relationships.map(({ type }) => type));
    expect(alignRelationshipKinds.every((types) => JSON.stringify(types) === JSON.stringify([
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
    ]))).toBe(true);

    const importedShapes = new Map(imported.slides.at(-1)!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    for (const fixture of fixtures) {
      const importedShape = importedShapes.get(fixture.name)!;
      expect(importedShape).toBeInstanceOf(ShapeModel);
      expect(importedShape.adjustments, fixture.name).toEqual(fixture.expected);
      expect(shapeXml(imported, alignTokens.length, importedShape.id), fixture.name)
        .toContain('<a:avLst>');
      expect(nativeSlide.addShape(fixture.type, {
        name: fixture.name,
        adjustments: fixture.expected,
      }).adjustments).toEqual(fixture.expected);
    }
    expect(nativeSlide.shapes.every((shape) =>
      !Object.hasOwn(shape, 'angleRange') && !Object.hasOwn(shape, 'arcThicknessRatio')))
      .toBe(true);
    expect(imported.diagnostics).toEqual([]);
    expect(native.diagnostics).toEqual([]);
  }, 20_000);

  it('closes PptxGenJS shape and text transform identity through strict native state', async () => {
    const classifications = [
      ...['flipH', 'flipV', 'objectName', 'rectRadius', 'rotate'].map((property) => ({
        id: `interface:ShapeProps@property:${property}`,
        status: 'deliberate-difference',
      })),
      {
        id: 'interface:ShapeProps@property:shapeName',
        status: 'defect-excluded',
      },
      ...['flipH', 'flipV', 'objectName', 'rectRadius', 'rotate'].map((property) => ({
        id: `interface:TextPropsOptions@property:${property}`,
        status: 'deliberate-difference',
      })),
      ...['isTextBox', 'shape'].map((property) => ({
        id: `interface:TextPropsOptions@property:${property}`,
        status: 'supported',
      })),
    ].sort((left, right) => left.id.localeCompare(right.id));
    expect(classifications).toHaveLength(13);
    expect(new Set(classifications.map(({ id }) => id)).size).toBe(13);
    expect({
      supported: classifications.filter(({ status }) => status === 'supported').length,
      deliberate: classifications.filter(
        ({ status }) => status === 'deliberate-difference',
      ).length,
      defects: classifications.filter(({ status }) => status === 'defect-excluded').length,
    }).toEqual({ supported: 2, deliberate: 10, defects: 1 });

    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    generatedSlide.addShape(generated.ShapeType.roundRect!, {
      x: 0.5,
      y: 4.5,
      w: 3,
      h: 1.5,
      rotate: 45,
      flipH: true,
      flipV: true,
      objectName: 'Canonical & object',
      rectRadius: 0.5,
      shapeName: 'Legacy & ignored',
    });
    generatedSlide.addShape(generated.ShapeType.rect!, {
      x: 8,
      y: 4.5,
      w: 1,
      h: 1,
      shapeName: 'Legacy only ignored',
    });
    generatedSlide.addText('Plain string overload', {
      x: 1,
      y: 1,
      w: 4,
      h: 2,
      rotate: 30,
      flipH: true,
      flipV: true,
      objectName: 'plain-wrapper',
      shape: 'roundRect',
      rectRadius: 0.25,
      isTextBox: true,
    });
    generatedSlide.addText([
      { text: 'Rich red', options: { bold: true, color: 'FF0000' } },
      { text: ' blue', options: { italic: true, color: '0000FF' } },
    ], {
      x: 1,
      y: 1,
      w: 5,
      h: 2,
      rotate: -15,
      flipV: true,
      objectName: 'rich-wrapper',
      shape: 'ellipse',
      isTextBox: false,
    });

    const imported = await openPptxGenJSPublicOutput(generated);
    const [importedShape, ignoredShape, importedPlain, importedRich] =
      imported.slides[0]!.shapes as readonly ShapeModel[];
    expect(importedShape).toBeInstanceOf(ShapeModel);
    expect({
      name: importedShape!.name,
      transform: importedShape!.transform,
      presetType: importedShape!.presetType,
      adjustments: importedShape!.adjustments,
    }).toEqual({
      name: 'Canonical & object',
      transform: {
        x: inches(0.5),
        y: inches(4.5),
        width: inches(3),
        height: inches(1.5),
        rotation: degrees(45),
        flipHorizontal: true,
        flipVertical: true,
      },
      presetType: 'roundRect',
      adjustments: [{ name: 'adj', value: 33_333 }],
    });
    expect(ignoredShape!.name).toBe('Shape 1');
    expect(shapeXml(imported, 0, ignoredShape!.id)).not.toContain('Legacy only ignored');
    expect({
      name: importedPlain!.name,
      text: importedPlain!.text,
      transform: importedPlain!.transform,
      presetType: importedPlain!.presetType,
      adjustments: importedPlain!.adjustments,
      isTextBox: importedPlain!.isTextBox,
    }).toEqual({
      name: 'plain-wrapper',
      text: 'Plain string overload',
      transform: {
        x: inches(1),
        y: inches(1),
        width: inches(4),
        height: inches(2),
        rotation: degrees(30),
        flipHorizontal: true,
        flipVertical: true,
      },
      presetType: 'roundRect',
      adjustments: [{ name: 'adj', value: 12_500 }],
      isTextBox: true,
    });
    expect({
      name: importedRich!.name,
      text: importedRich!.text,
      rotation: importedRich!.transform.rotation,
      flipVertical: importedRich!.transform.flipVertical,
      presetType: importedRich!.presetType,
      isTextBox: importedRich!.isTextBox,
    }).toEqual({
      name: 'rich-wrapper',
      text: 'Rich red blue',
      rotation: degrees(-15),
      flipVertical: true,
      presetType: 'ellipse',
      isTextBox: false,
    });

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const nativeShape = nativeSlide.addShape('roundRect', {
      name: 'Canonical & object',
      x: inches(0.5),
      y: inches(4.5),
      width: inches(3),
      height: inches(1.5),
      rotation: degrees(45),
      flipHorizontal: true,
      flipVertical: true,
      adjustments: [{ name: 'adj', value: 33_333 }],
    });
    const nativePlain = nativeSlide.addText('Plain string overload', {
      name: 'plain-wrapper',
      x: inches(1),
      y: inches(1),
      width: inches(4),
      height: inches(2),
      rotation: degrees(30),
      flipHorizontal: true,
      flipVertical: true,
      shape: 'roundRect',
      rectRadius: inches(0.25),
      isTextBox: true,
    });
    expect(nativeShape.transform).toEqual(importedShape!.transform);
    expect(nativeShape.adjustments).toEqual(importedShape!.adjustments);
    expect(nativePlain.transform).toEqual(importedPlain!.transform);
    expect(nativePlain.adjustments).toEqual(importedPlain!.adjustments);
    nativeShape.name = 'Strict shape edited';
    nativePlain.name = '';
    const reopenedNative = await PptxDocument.open(await native.write());
    expect(reopenedNative.slides[0]!.shapes.map(({ name }) => name))
      .toEqual(['Strict shape edited', '']);
  }, 20_000);

  it('matches representative preset shape public output semantically', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const cases: readonly {
      readonly type: PresetShapeType;
      readonly generatedOptions?: Record<string, unknown>;
      readonly nativeOptions?: AddShapeOptions;
      readonly expectedGeneratedName: string;
    }[] = [
      { type: 'rect', expectedGeneratedName: 'Shape 0' },
      { type: 'ellipse', generatedOptions: {}, nativeOptions: {}, expectedGeneratedName: 'Shape 0' },
      { type: 'line', expectedGeneratedName: 'Shape 0' },
      { type: 'lineInv', expectedGeneratedName: 'Shape 0' },
      { type: 'flowChartDecision', expectedGeneratedName: 'Shape 0' },
      { type: 'star5', expectedGeneratedName: 'Shape 0' },
      { type: 'actionButtonHome', expectedGeneratedName: 'Shape 0' },
      {
        type: 'roundRect',
        generatedOptions: {
          objectName: 'Public shape',
          x: 1.25,
          y: 2.5,
          w: 3.75,
          h: 4.5,
          rotate: 45,
          flipH: true,
          flipV: true,
        },
        nativeOptions: {
          name: 'Public shape',
          x: inches(1.25),
          y: inches(2.5),
          width: inches(3.75),
          height: inches(4.5),
          rotation: degrees(45),
          flipHorizontal: true,
          flipVertical: true,
        },
        expectedGeneratedName: 'Public shape',
      },
    ];

    for (const fixture of cases) {
      const publicType = generated.ShapeType[fixture.type];
      expect(publicType, fixture.type).toBe(fixture.type);
      const slide = generated.addSlide();
      if (Object.prototype.hasOwnProperty.call(fixture, 'generatedOptions')) {
        slide.addShape(publicType!, fixture.generatedOptions);
      } else {
        slide.addShape(publicType!);
      }
    }

    const imported = await openPptxGenJSPublicOutput(generated);
    const native = PptxDocument.create();
    for (const [index, fixture] of cases.entries()) {
      const importedShape = imported.slides[index]?.shapes[0];
      const nativeShape = native.addSlide().addShape(fixture.type, fixture.nativeOptions);
      expect(importedShape, fixture.type).toBeInstanceOf(ShapeModel);
      expect(importedShape?.name, fixture.type).toBe(fixture.expectedGeneratedName);
      expect((importedShape as ShapeModel).presetType, fixture.type).toBe(fixture.type);
      expect(importedShape?.transform, fixture.type).toEqual(nativeShape.transform);
      expect(directShapePaintState(shapeXml(imported, index, importedShape!.id)))
        .toEqual(directShapePaintState(shapeXml(native, index, nativeShape.id)));
    }
  });

  it('compares representative text shape preset geometry public output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const slide = generated.addSlide();
    const cases: readonly {
      readonly name: string;
      readonly text: Parameters<PptxGenJSSlide['addText']>[0];
      readonly expectedText: string;
      readonly type: PresetShapeType;
      readonly options: Readonly<Record<string, unknown>>;
    }[] = [
      {
        name: 'Geometry omitted',
        text: 'Plain omitted',
        expectedText: 'Plain omitted',
        type: 'rect',
        options: {},
      },
      {
        name: 'Geometry undefined',
        text: 'Plain undefined',
        expectedText: 'Plain undefined',
        type: 'rect',
        options: { shape: undefined },
      },
      {
        name: 'Geometry ellipse',
        text: 'Ellipse',
        expectedText: 'Ellipse',
        type: 'ellipse',
        options: { shape: 'ellipse' },
      },
      {
        name: 'Geometry round rectangle',
        text: 'Round rectangle',
        expectedText: 'Round rectangle',
        type: 'roundRect',
        options: { shape: 'roundRect' },
      },
      {
        name: 'Geometry line',
        text: 'Line',
        expectedText: 'Line',
        type: 'line',
        options: { shape: 'line', line: {} },
      },
      {
        name: 'Geometry inverse line',
        text: 'Inverse line',
        expectedText: 'Inverse line',
        type: 'lineInv',
        options: { shape: 'lineInv' },
      },
      {
        name: 'Geometry flowchart',
        text: 'First\nSecond',
        expectedText: 'First\nSecond',
        type: 'flowChartDecision',
        options: { shape: 'flowChartDecision' },
      },
      {
        name: 'Geometry callout',
        text: 'Callout',
        expectedText: 'Callout',
        type: 'wedgeRoundRectCallout',
        options: { shape: 'wedgeRoundRectCallout' },
      },
      {
        name: 'Geometry action button',
        text: 'Action button',
        expectedText: 'Action button',
        type: 'actionButtonHome',
        options: { shape: 'actionButtonHome' },
      },
      {
        name: 'Geometry star',
        text: [
          { text: 'Rich ', options: { bold: true } },
          { text: 'star', options: { color: 'FF0000' } },
        ],
        expectedText: 'Rich star',
        type: 'star5',
        options: { shape: 'star5' },
      },
    ];

    for (const [index, fixture] of cases.entries()) {
      slide.addText(fixture.text, {
        objectName: fixture.name,
        x: 1,
        y: 0.25 + index * 0.25,
        w: 4,
        h: 0.2,
        ...fixture.options,
      });
    }

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShapes = imported.slides[0]!.shapes as readonly ShapeModel[];
    expect(importedShapes.map(({ name }) => name)).toEqual(cases.map(({ name }) => name));
    for (const [index, fixture] of cases.entries()) {
      const shape = importedShapes[index]!;
      expect(shape, fixture.name).toBeInstanceOf(ShapeModel);
      expect(shape.text, fixture.name).toBe(fixture.expectedText);
      expect(shape.presetType, fixture.name).toBe(fixture.type);
      expect(directTextPresetGeometryState(shapeXml(imported, 0, shape.id)), fixture.name)
        .toEqual({ type: fixture.type, adjustments: [] });
    }
  });

  it('compares every common text shape preset geometry token', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const commonTypes = PRESET_SHAPE_TYPES.filter((type) => type !== 'foldedCorner');
    const generatedSlide = generated.addSlide();
    for (const type of commonTypes) {
      const publicType = generated.ShapeType[type];
      expect(publicType, type).toBe(type);
      generatedSlide.addText(type, {
        objectName: type,
        shape: publicType,
        ...(type === 'line' ? { line: {} } : {}),
      });
    }

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShapes = imported.slides[0]!.shapes as readonly ShapeModel[];
    expect(importedShapes).toHaveLength(177);
    expect(importedShapes.map(({ presetType }) => presetType)).toEqual(commonTypes);
    expect(importedShapes.map(({ text }) => text)).toEqual(commonTypes);

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const nativeShapes = commonTypes.map((type) => nativeSlide.addText(type, {
      name: type,
      shape: type,
    }));
    expect(nativeShapes.map(({ presetType }) => presetType)).toEqual(commonTypes);
    for (const [index, type] of commonTypes.entries()) {
      const importedShape = importedShapes[index]!;
      const nativeShape = nativeShapes[index]!;
      expect(directTextPresetGeometryState(shapeXml(imported, 0, importedShape.id)), type)
        .toEqual(directTextPresetGeometryState(shapeXml(native, 0, nativeShape.id)));
    }
  }, 30_000);

  it('locks text shape preset geometry upstream defects and native strictness', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    expect(generated.ShapeType.folderCorner).toBe('folderCorner');
    expect(generated.ShapeType.foldedCorner).toBeUndefined();
    expect(generated.ShapeType.custGeom).toBe('custGeom');
    const generatedSlide = generated.addSlide();
    const divergentCases: readonly {
      readonly name: string;
      readonly shape: unknown;
      readonly expectedType?: string;
      readonly points?: readonly PptxGenJSCustomPoint[];
    }[] = [
      { name: 'Falsy false', shape: false, expectedType: 'rect' },
      { name: 'Falsy empty', shape: '', expectedType: 'rect' },
      { name: 'Unknown string', shape: 'unknown', expectedType: 'unknown' },
      { name: 'Numeric token', shape: 42, expectedType: '42' },
      { name: 'Folder corner', shape: 'folderCorner', expectedType: 'folderCorner' },
      {
        name: 'Custom geometry',
        shape: 'custGeom',
        points: [{ x: 0, y: 0, moveTo: true }, { x: 1, y: 1 }],
      },
    ];
    for (const fixture of divergentCases) {
      generatedSlide.addText(fixture.name, {
        objectName: fixture.name,
        x: 1,
        y: 1,
        w: 2,
        h: 1,
        shape: fixture.shape,
        ...(fixture.points === undefined ? {} : { points: fixture.points }),
      });
    }

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    for (const fixture of divergentCases) {
      const shape = importedShapes.get(fixture.name)!;
      if (fixture.expectedType === 'rect') {
        expect(shape.presetType, fixture.name).toBe('rect');
      } else {
        expect(shape.presetType, fixture.name).toBeUndefined();
      }
      if (fixture.expectedType !== undefined) {
        expect(directTextPresetGeometryState(shapeXml(imported, 0, shape.id)), fixture.name)
          .toEqual({ type: fixture.expectedType, adjustments: [] });
      }
    }
    const custom = importedShapes.get('Custom geometry')!;
    expect(custom.customGeometry).toBeDefined();
    expect(shapeXml(imported, 0, custom.id)).toContain('<a:custGeom>');

    const singlePublicShape = async (options: Readonly<Record<string, unknown>>) => {
      const presentation = new PptxGenJS();
      const presentationSlide = presentation.addSlide();
      presentationSlide.addText('Isolation', {
        objectName: 'Isolation',
        x: 1,
        y: 1,
        w: 2,
        h: 1,
        shape: 'roundRect',
        ...options,
      });
      const document = await openPptxGenJSPublicOutput(presentation);
      const shape = document.slides[0]!.shapes[0] as ShapeModel;
      return shapeXml(document, 0, shape.id);
    };
    const textBoxFalse = await singlePublicShape({ isTextBox: false });
    const textBoxTrue = await singlePublicShape({ isTextBox: true });
    expect(textBoxFalse).toContain('<p:cNvSpPr/>');
    expect(textBoxTrue).toContain('<p:cNvSpPr txBox="1"/>');
    expect(textBoxTrue.replace('<p:cNvSpPr txBox="1"/>', '<p:cNvSpPr/>'))
      .toBe(textBoxFalse);

    const zeroRadius = await singlePublicShape({ rectRadius: 0 });
    const positiveRadius = await singlePublicShape({ rectRadius: 0.5 });
    expect(directTextPresetGeometryState(zeroRadius))
      .toEqual({ type: 'roundRect', adjustments: [] });
    expect(directTextPresetGeometryState(positiveRadius)).toEqual({
      type: 'roundRect',
      adjustments: [{ name: 'adj', formula: 'val 50000' }],
    });
    expect(positiveRadius.replace(
      '<a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst>',
      '<a:avLst></a:avLst>',
    )).toBe(zeroRadius);

    const brokenLine = new PptxGenJS();
    expect(() => brokenLine.addSlide().addText('Broken line', { shape: 'line' }))
      .toThrow(TypeError);

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const existing = nativeSlide.addText('Existing', { name: 'Existing' });
    const beforeInvalid = packageState(native);
    const beforeShapes = nativeSlide.shapes;
    for (const shape of [false, '', 'unknown', 42, 'folderCorner', 'custGeom']) {
      expect(() => nativeSlide.addText('Invalid geometry', { shape } as never))
        .toThrow(TypeError);
      expect(packageState(native)).toEqual(beforeInvalid);
      expect(nativeSlide.shapes).toEqual(beforeShapes);
      expect(nativeSlide.shapes[0]).toBe(existing);
    }

    const nativeLine = nativeSlide.addText('Native line', {
      name: 'Native line',
      shape: 'line',
    });
    expect(nativeLine.presetType).toBe('line');
    expect(shapeXml(native, 0, nativeLine.id)).toContain(
      '</a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>',
    );
    const foldedCorner = nativeSlide.addText('Folded corner', {
      name: 'Folded corner',
      shape: 'foldedCorner',
    });
    expect(foldedCorner.presetType).toBe('foldedCorner');
    const reopened = await PptxDocument.open(await native.write());
    expect((reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'Folded corner',
    ) as ShapeModel).presetType).toBe('foldedCorner');
  });

  it('compares text box state public output across text and placeholder owners', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const publicObjects: PptxGenJSMasterObject[] = [
      {
        text: {
          text: 'Layout shape text',
          options: { objectName: 'layout-shape-text', x: 0.5, y: 0.5, w: 2, h: 0.5 },
        },
      },
      {
        text: {
          text: 'Layout text box',
          options: {
            objectName: 'layout-text-box', x: 3, y: 0.5, w: 2, h: 0.5,
            isTextBox: true,
          },
        },
      },
      ...[
        {
          name: 'populate_false_source', objectName: 'populate-false-source',
          type: 'title', y: 1.5, isTextBox: false,
        },
        {
          name: 'populate_true_source', objectName: 'populate-true-source',
          type: 'body', y: 2.5, isTextBox: true,
        },
        {
          name: 'empty_false_source', objectName: 'empty-false-source',
          type: 'body', y: 3.5, isTextBox: false,
        },
        {
          name: 'empty_true_source', objectName: 'empty-true-source',
          type: 'body', y: 4.5, isTextBox: true,
        },
      ].map((options, index): PptxGenJSMasterObject => ({
        placeholder: {
          text: `Prompt ${index}`,
          options: {
            ...options,
            x: 0.5,
            w: 4,
            h: 0.5,
          } as PptxGenJSPlaceholderOptions,
        },
      })),
    ];
    generated.defineSlideMaster({ title: 'TEXT-BOX-STATE', objects: publicObjects });
    const generatedSlide = generated.addSlide({ masterName: 'TEXT-BOX-STATE' });
    generatedSlide.addText('Omitted', {
      objectName: 'slide-omitted', x: 5, y: 0.5, w: 2, h: 0.5,
    });
    generatedSlide.addText('Undefined', {
      objectName: 'slide-undefined', x: 5, y: 1.5, w: 2, h: 0.5,
      isTextBox: undefined,
    });
    generatedSlide.addText('False', {
      objectName: 'slide-false', x: 5, y: 2.5, w: 2, h: 0.5,
      isTextBox: false,
    });
    generatedSlide.addText('True', {
      objectName: 'slide-true', x: 5, y: 3.5, w: 2, h: 0.5,
      isTextBox: true,
    });
    generatedSlide.addText([
      { text: 'Rich', options: { bold: true } },
      { text: ' text box', options: { italic: true } },
    ], {
      objectName: 'slide-rich-true', x: 5, y: 4.5, w: 2, h: 0.5,
      isTextBox: true,
    });
    generatedSlide.addText('Population true', {
      placeholder: 'populate_false_source',
      isTextBox: true,
    });
    generatedSlide.addText('Population false', {
      placeholder: 'populate_true_source',
    });

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedLayout = imported.layouts.find(({ name }) => name === 'TEXT-BOX-STATE')!;
    const states = (shapes: readonly ShapeModel[]) => Object.fromEntries(
      shapes.map((shape) => [shape.name, shape.isTextBox]),
    );
    expect(states(importedLayout.shapes as readonly ShapeModel[])).toMatchObject({
      'layout-shape-text': false,
      'layout-text-box': true,
      'populate-false-source': false,
      'populate-true-source': true,
      'empty-false-source': false,
      'empty-true-source': true,
    });
    expect(states(imported.slides[0]!.shapes as readonly ShapeModel[])).toMatchObject({
      'slide-omitted': false,
      'slide-undefined': false,
      'slide-false': false,
      'slide-true': true,
      'slide-rich-true': true,
      'populate-false-source': false,
      'populate-true-source': true,
      'empty-false-source': false,
      'empty-true-source': true,
    });

    const native = PptxDocument.create();
    const nativeLayout = await native.defineSlideMaster({
      title: 'TEXT-BOX-STATE-NATIVE',
      objects: [
        {
          kind: 'text', text: 'Layout shape text',
          options: { name: 'layout-shape-text', isTextBox: false },
        },
        {
          kind: 'text', text: 'Layout text box',
          options: { name: 'layout-text-box', isTextBox: true },
        },
        {
          kind: 'placeholder', text: 'Prompt 0',
          options: {
            name: 'populate-false-source', type: 'title', index: 100,
            isTextBox: false,
          },
        },
        {
          kind: 'placeholder', text: 'Prompt 1',
          options: {
            name: 'populate-true-source', type: 'body', index: 101,
            isTextBox: true,
          },
        },
        {
          kind: 'placeholder', text: 'Prompt 2',
          options: {
            name: 'empty-false-source', type: 'body', index: 102,
            isTextBox: false,
          },
        },
        {
          kind: 'placeholder', text: 'Prompt 3',
          options: {
            name: 'empty-true-source', type: 'body', index: 103,
            isTextBox: true,
          },
        },
      ],
    });
    const nativeSlide = native.addSlide({ masterName: nativeLayout.name });
    nativeSlide.addText('Omitted', { name: 'slide-omitted' });
    nativeSlide.addText('Undefined', { name: 'slide-undefined', isTextBox: undefined } as never);
    nativeSlide.addText('False', { name: 'slide-false', isTextBox: false });
    nativeSlide.addText('True', { name: 'slide-true', isTextBox: true });
    nativeSlide.addRichText([{ runs: [{ text: 'Rich', style: { bold: true } }] }], {
      name: 'slide-rich-true',
      isTextBox: true,
    });
    nativeSlide.addText('Population true', {
      placeholder: 'populate-false-source',
      isTextBox: true,
    });
    nativeSlide.addText('Population false', { placeholder: 'populate-true-source' });
    expect(states(nativeLayout.shapes as readonly ShapeModel[])).toEqual(
      states(importedLayout.shapes as readonly ShapeModel[]),
    );
    expect(states(nativeSlide.shapes as readonly ShapeModel[])).toEqual(
      states(imported.slides[0]!.shapes as readonly ShapeModel[]),
    );

    const styledPublicXml = async (isTextBox: boolean): Promise<string> => {
      const presentation = new PptxGenJS();
      const slide = presentation.addSlide();
      slide.addText([
        { text: 'Styled', options: { bold: true, color: '336699' } },
        { text: ' text', options: { italic: true } },
      ], {
        objectName: 'Styled isolation',
        x: 1,
        y: 1,
        w: 4,
        h: 2,
        shape: 'roundRect',
        rectRadius: 0.5,
        fill: { color: 'ABCDEF' },
        line: { color: '123456', width: 2, dash: 'dashDot' },
        hyperlink: { url: 'https://example.com', tooltip: 'Keep' },
        margin: 0,
        valign: 'bottom',
        isTextBox,
      });
      const document = await openPptxGenJSPublicOutput(presentation);
      return shapeXml(document, 0, document.slides[0]!.shapes[0]!.id);
    };
    const publicFalseXml = await styledPublicXml(false);
    const publicTrueXml = await styledPublicXml(true);
    expect(publicTrueXml.replace('<p:cNvSpPr txBox="1"/>', '<p:cNvSpPr/>'))
      .toBe(publicFalseXml);

    const styledNativeXml = (isTextBox: boolean): string => {
      const presentation = PptxDocument.create();
      const shape = presentation.addSlide().addRichText([{
        runs: [
          { text: 'Styled', style: { bold: true, color: { kind: 'srgb', value: '336699' } } },
          { text: ' text', style: { italic: true } },
        ],
      }], {
        name: 'Styled isolation',
        shape: 'roundRect',
        rectRadius: inches(0.5),
        width: inches(4),
        height: inches(2),
        fill: { kind: 'solid', color: { kind: 'srgb', value: 'ABCDEF' } },
        line: {
          kind: 'line', color: { kind: 'srgb', value: '123456' },
          width: 2, dash: 'dashDot',
        },
        hyperlink: { url: 'https://example.com', tooltip: 'Keep' },
        margin: 0,
        valign: 'bottom',
        isTextBox,
      });
      return shapeXml(presentation, 0, shape.id);
    };
    const nativeFalseXml = styledNativeXml(false);
    const nativeTrueXml = styledNativeXml(true);
    expect(nativeTrueXml.replace('<p:cNvSpPr txBox="1"/>', '<p:cNvSpPr/>'))
      .toBe(nativeFalseXml);
  });

  it('locks text box runtime truthiness divergence and native strictness', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const cases = [
      { name: 'Truthy string', value: 'yes', expected: true },
      { name: 'Truthy number', value: 1, expected: true },
      { name: 'Truthy object', value: {}, expected: true },
      { name: 'Truthy array', value: [], expected: true },
      { name: 'Falsy zero', value: 0, expected: false },
      { name: 'Falsy empty', value: '', expected: false },
      { name: 'Falsy null', value: null, expected: false },
    ] as const;
    for (const [index, fixture] of cases.entries()) {
      generatedSlide.addText(fixture.name, {
        objectName: fixture.name,
        x: 0.5,
        y: 0.5 + index * 0.5,
        w: 2,
        h: 0.4,
        isTextBox: fixture.value,
      });
    }
    const imported = await openPptxGenJSPublicOutput(generated);
    expect((imported.slides[0]!.shapes as readonly ShapeModel[]).map(
      ({ name, isTextBox }) => ({ name, isTextBox }),
    )).toEqual(cases.map(({ name, expected }) => ({ name, isTextBox: expected })));

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const existing = nativeSlide.addText('Existing', { name: 'Existing', isTextBox: true });
    const before = packageState(native);
    const shapes = nativeSlide.shapes;
    for (const fixture of cases) {
      expect(() => nativeSlide.addText(fixture.name, {
        isTextBox: fixture.value,
      } as never), fixture.name).toThrow(TypeError);
      expect(packageState(native), fixture.name).toEqual(before);
      expect(nativeSlide.shapes, fixture.name).toEqual(shapes);
      expect(nativeSlide.shapes[0], fixture.name).toBe(existing);
    }
  });

  it('compares public rich text line breaks with native canonical paragraphs', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const addPublic = (
      name: string,
      runs: readonly { readonly text: string; readonly options?: Record<string, unknown> }[],
      y: number,
    ) => generatedSlide.addText(runs, {
      objectName: name,
      x: 0.5,
      y,
      w: 4,
      h: 0.5,
    });
    addPublic('break-middle', [
      { text: 'A', options: { breakLine: true } },
      { text: 'B', options: {} },
    ], 0.5);
    addPublic('break-trailing', [
      { text: 'A', options: { breakLine: true } },
    ], 1);
    addPublic('break-empty', [
      { text: '', options: { breakLine: true } },
      { text: 'B', options: {} },
    ], 1.5);
    addPublic('break-consecutive-empty', [
      { text: '', options: { breakLine: true } },
      { text: '', options: { breakLine: true } },
      { text: 'C', options: {} },
    ], 2);
    addPublic('break-false', [
      { text: 'A', options: { breakLine: false } },
      { text: 'B', options: {} },
    ], 2.5);
    addPublic('break-soft', [
      { text: 'A', options: { breakLine: true } },
      { text: 'B', options: { softBreakBefore: true } },
      { text: 'C', options: {} },
    ], 3);
    addPublic('break-properties', [
      {
        text: 'Right',
        options: {
          breakLine: true,
          align: 'right',
          paraSpaceAfter: 8,
          tabStops: [{ position: 1.5, alignment: 'r' }],
        },
      },
      {
        text: 'Center',
        options: {
          align: 'center',
          paraSpaceAfter: 4,
          tabStops: [{ position: 2.5, alignment: 'ctr' }],
        },
      },
    ], 3.5);
    addPublic('break-hyperlinks', [
      {
        text: 'External',
        options: {
          breakLine: true,
          hyperlink: { url: 'https://break.example', tooltip: 'External' },
        },
      },
      {
        text: 'Internal',
        options: { hyperlink: { slide: 2, tooltip: '' } },
      },
    ], 4);
    generated.addSlide();

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShapes = imported.slides[0]!.shapes as readonly ShapeModel[];
    const byName = (name: string) => importedShapes.find((shape) => shape.name === name)!;
    const paragraphTexts = (shape: ShapeModel) => shape.richText.map(({ runs }) =>
      runs.map(({ text }) => text));
    const expected = new Map<string, readonly (readonly string[])[]>([
      ['break-middle', [['A'], ['B']]],
      ['break-trailing', [['A']]],
      ['break-empty', [[], ['B']]],
      ['break-consecutive-empty', [[], [], ['C']]],
      ['break-false', [['A', 'B']]],
      ['break-soft', [['A'], ['B', 'C']]],
      ['break-properties', [['Right'], ['Center']]],
      ['break-hyperlinks', [['External'], ['Internal']]],
    ]);
    for (const [name, paragraphs] of expected) {
      expect(paragraphTexts(byName(name)), name).toEqual(paragraphs);
      expect(shapeXml(imported, 0, byName(name).id).match(/<a:p>/g), name)
        .toHaveLength(paragraphs.length);
    }
    expect(byName('break-properties').richText.map(({ align, spacing, tabStops }) => ({
      align,
      spacing,
      tabStops,
    }))).toEqual([
      {
        align: 'right',
        spacing: { after: 8 },
        tabStops: [{ position: 1.5, alignment: 'right' }],
      },
      {
        align: 'center',
        spacing: { after: 4 },
        tabStops: [{ position: 2.5, alignment: 'center' }],
      },
    ]);
    expect(byName('break-hyperlinks').richText.map(({ runs }) =>
      runs[0]?.style?.hyperlink)).toEqual([
      { url: 'https://break.example', tooltip: 'External' },
      { slide: 2, tooltip: '' },
    ]);
    expect(byName('break-soft').richText[1]!.runs[0]!.softBreakBefore).toBeUndefined();

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    native.addSlide();
    const nativeFixtures = [
      {
        name: 'break-middle',
        paragraphs: [{ runs: [
          { text: 'A', breakLine: true },
          { text: 'B' },
        ] }],
      },
      {
        name: 'break-trailing',
        paragraphs: [{ runs: [{ text: 'A', breakLine: true }] }],
      },
      {
        name: 'break-empty',
        paragraphs: [{ runs: [
          { text: '', breakLine: true },
          { text: 'B' },
        ] }],
      },
      {
        name: 'break-consecutive-empty',
        paragraphs: [{ runs: [
          { text: '', breakLine: true },
          { text: '', breakLine: true },
          { text: 'C' },
        ] }],
      },
      {
        name: 'break-false',
        paragraphs: [{ runs: [
          { text: 'A', breakLine: false },
          { text: 'B' },
        ] }],
      },
      {
        name: 'break-soft',
        paragraphs: [{ runs: [
          { text: 'A', breakLine: true },
          { text: 'B', softBreakBefore: true },
          { text: 'C' },
        ] }],
      },
      {
        name: 'break-hyperlinks',
        paragraphs: [{ runs: [
          {
            text: 'External',
            breakLine: true,
            style: { hyperlink: { url: 'https://break.example', tooltip: 'External' } },
          },
          {
            text: 'Internal',
            style: { hyperlink: { slide: 2, tooltip: '' } },
          },
        ] }],
      },
    ] as const;
    for (const fixture of nativeFixtures) {
      const shape = nativeSlide.addRichText(fixture.paragraphs, { name: fixture.name });
      expect(paragraphTexts(shape), fixture.name).toEqual(expected.get(fixture.name));
    }
    const nativeProperties = nativeSlide.addRichText([
      {
        align: 'right',
        spacing: { after: 8 },
        tabStops: [{ position: 1.5, alignment: 'right' }],
        runs: [{ text: 'Right' }],
      },
      {
        align: 'center',
        spacing: { after: 4 },
        tabStops: [{ position: 2.5, alignment: 'center' }],
        runs: [{ text: 'Center' }],
      },
    ], { name: 'break-properties' });
    expect(nativeProperties.richText.map(({ align, spacing, tabStops }) => ({
      align,
      spacing,
      tabStops,
    }))).toEqual(byName('break-properties').richText.map(({ align, spacing, tabStops }) => ({
      align,
      spacing,
      tabStops,
    })));
    const nativeHyperlinks = nativeSlide.shapes.find(
      ({ name }) => name === 'break-hyperlinks',
    ) as ShapeModel;
    expect(nativeHyperlinks.richText.map(({ runs }) => runs[0]?.style?.hyperlink))
      .toEqual(byName('break-hyperlinks').richText.map(
        ({ runs }) => runs[0]?.style?.hyperlink,
      ));
    const nativeSoft = nativeSlide.shapes.find(({ name }) => name === 'break-soft') as ShapeModel;
    expect(nativeSoft.richText[1]!.runs[0]!.softBreakBefore).toBe(true);
    expect((await imported.write()).byteLength).toBeGreaterThan(0);
    expect((await native.write()).byteLength).toBeGreaterThan(0);
  });

  it('locks rich text line break runtime truthiness divergence and native strictness', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const cases = [
      { name: 'Truthy string', value: 'yes', paragraphs: 2 },
      { name: 'Truthy number', value: 1, paragraphs: 2 },
      { name: 'Truthy object', value: {}, paragraphs: 2 },
      { name: 'Truthy array', value: [], paragraphs: 2 },
      { name: 'Falsy zero', value: 0, paragraphs: 1 },
      { name: 'Falsy empty', value: '', paragraphs: 1 },
      { name: 'Falsy null', value: null, paragraphs: 1 },
      { name: 'Falsy undefined', value: undefined, paragraphs: 1 },
    ] as const;
    for (const [index, fixture] of cases.entries()) {
      generatedSlide.addText([
        { text: 'A', options: { breakLine: fixture.value } },
        { text: 'B', options: {} },
      ], {
        objectName: fixture.name,
        x: 0.5,
        y: 0.5 + index * 0.5,
        w: 3,
        h: 0.4,
      });
    }
    const imported = await openPptxGenJSPublicOutput(generated);
    for (const fixture of cases) {
      const shape = imported.slides[0]!.shapes.find(
        ({ name }) => name === fixture.name,
      ) as ShapeModel;
      expect(shape.richText, fixture.name).toHaveLength(fixture.paragraphs);
    }

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const existing = nativeSlide.addRichText([{ runs: [
      { text: 'A', breakLine: undefined },
      { text: 'B' },
    ] }] as never, { name: 'Falsy undefined' });
    expect(existing.richText).toHaveLength(1);
    const before = packageState(native);
    const shapes = nativeSlide.shapes;
    for (const fixture of cases.slice(0, -1)) {
      expect(() => nativeSlide.addRichText([{ runs: [
        { text: 'A', breakLine: fixture.value },
        { text: 'B' },
      ] }] as never, { name: fixture.name }), fixture.name).toThrow(TypeError);
      expect(packageState(native), fixture.name).toEqual(before);
      expect(nativeSlide.shapes, fixture.name).toEqual(shapes);
      expect(nativeSlide.shapes[0], fixture.name).toBe(existing);
    }
  });

  it('compares supported text shape rectangle radius public output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const fixtures = [
      {
        name: 'Radius two by one',
        radius: 0.5,
        width: 2,
        height: 1,
        expected: 50_000,
      },
      {
        name: 'Radius four by two',
        radius: 0.5,
        width: 4,
        height: 2,
        expected: 25_000,
      },
      {
        name: 'Radius one inch',
        radius: 1,
        width: 4,
        height: 2,
        expected: 50_000,
      },
      {
        name: 'Radius fractional',
        radius: 0.333333,
        width: 3,
        height: 1.5,
        expected: 22_222,
      },
    ] as const;

    for (const [index, fixture] of fixtures.entries()) {
      generatedSlide.addText(fixture.name, {
        objectName: fixture.name,
        x: 0.5,
        y: 0.5 + index,
        w: fixture.width,
        h: fixture.height,
        shape: generated.ShapeType.roundRect,
        rectRadius: fixture.radius,
      });
    }

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedByName = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    for (const fixture of fixtures) {
      const importedShape = importedByName.get(fixture.name)!;
      const nativeShape = nativeSlide.addText(fixture.name, {
        name: fixture.name,
        shape: 'roundRect',
        rectRadius: inches(fixture.radius),
        width: inches(fixture.width),
        height: inches(fixture.height),
      });
      const expected = [{ name: 'adj', value: fixture.expected }];
      expect(importedShape.presetType, fixture.name).toBe('roundRect');
      expect(importedShape.text, fixture.name).toBe(fixture.name);
      expect(importedShape.adjustments, fixture.name).toEqual(expected);
      expect(nativeShape.presetType, fixture.name).toBe('roundRect');
      expect(nativeShape.text, fixture.name).toBe(fixture.name);
      expect(nativeShape.adjustments, fixture.name).toEqual(expected);
      expect(directTextPresetGeometryState(
        shapeXml(imported, 0, importedShape.id),
      ), fixture.name).toEqual({
        type: 'roundRect',
        adjustments: [{ name: 'adj', formula: `val ${fixture.expected}` }],
      });
      expect(directTextPresetGeometryState(shapeXml(native, 0, nativeShape.id)), fixture.name)
        .toEqual({
          type: 'roundRect',
          adjustments: [{ name: 'adj', formula: `val ${fixture.expected}` }],
        });
    }

    const publicShapeXml = async (rectRadius?: number) => {
      const presentation = new PptxGenJS();
      presentation.addSlide().addText('Radius isolation', {
        objectName: 'Radius isolation',
        x: 1,
        y: 1,
        w: 4,
        h: 2,
        shape: presentation.ShapeType.roundRect,
        ...(rectRadius === undefined ? {} : { rectRadius }),
      });
      const document = await openPptxGenJSPublicOutput(presentation);
      return shapeXml(document, 0, document.slides[0]!.shapes[0]!.id);
    };
    const publicOmitted = await publicShapeXml();
    const publicPositive = await publicShapeXml(0.5);
    expect(publicPositive.replace(
      '<a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst>',
      '<a:avLst></a:avLst>',
    )).toBe(publicOmitted);

    const nativeShapeXml = (withRadius: boolean) => {
      const document = PptxDocument.create();
      const shape = document.addSlide().addRichText([{
        runs: [
          { text: 'Rounded', style: { hyperlink: { url: 'https://run.example' } } },
          { text: ' isolation' },
        ],
      }], {
        name: 'Radius isolation',
        shape: 'roundRect',
        ...(withRadius ? { rectRadius: inches(0.5) } : {}),
        width: inches(4),
        height: inches(2),
        fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' } },
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '123ABC' },
          width: 2,
          dash: 'dashDot',
        },
        arrows: { begin: 'oval', end: 'triangle' },
        shadow: { kind: 'outer', opacity: 0.5 },
        hyperlink: { url: 'https://shape.example', tooltip: 'Shape' },
        margin: 0,
        valign: 'bottom',
        vert: 'vert',
        fit: 'shrink',
        wrap: false,
      });
      return shapeXml(document, 0, shape.id);
    };
    const nativeOmitted = nativeShapeXml(false);
    const nativePositive = nativeShapeXml(true);
    expect(nativePositive.replace(
      '<a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst>',
      '<a:avLst/>',
    )).toBe(nativeOmitted);
  });

  it('locks text shape rectangle radius upstream divergences and native strictness', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const generatedCases: readonly {
      readonly name: string;
      readonly shape: string;
      readonly radius: unknown;
      readonly expectedType: string;
      readonly expectedFormula?: string;
    }[] = [
      {
        name: 'Dropped zero radius',
        shape: 'roundRect',
        radius: 0,
        expectedType: 'roundRect',
      },
      {
        name: 'String radius coercion',
        shape: 'roundRect',
        radius: '0.5',
        expectedType: 'roundRect',
        expectedFormula: 'val 50000',
      },
      {
        name: 'Negative radius passthrough',
        shape: 'roundRect',
        radius: -0.25,
        expectedType: 'roundRect',
        expectedFormula: 'val -25000',
      },
      {
        name: 'Over range radius passthrough',
        shape: 'roundRect',
        radius: 1.5,
        expectedType: 'roundRect',
        expectedFormula: 'val 150000',
      },
      {
        name: 'Dropped NaN radius',
        shape: 'roundRect',
        radius: Number.NaN,
        expectedType: 'roundRect',
      },
      {
        name: 'Infinite radius formula',
        shape: 'roundRect',
        radius: Number.POSITIVE_INFINITY,
        expectedType: 'roundRect',
        expectedFormula: 'val Infinity',
      },
      {
        name: 'Negative infinite radius formula',
        shape: 'roundRect',
        radius: Number.NEGATIVE_INFINITY,
        expectedType: 'roundRect',
        expectedFormula: 'val -Infinity',
      },
      {
        name: 'Rectangle radius passthrough',
        shape: 'rect',
        radius: 0.5,
        expectedType: 'rect',
        expectedFormula: 'val 50000',
      },
      {
        name: 'Ellipse radius passthrough',
        shape: 'ellipse',
        radius: 0.5,
        expectedType: 'ellipse',
        expectedFormula: 'val 50000',
      },
    ];

    for (const [index, fixture] of generatedCases.entries()) {
      generatedSlide.addText(fixture.name, {
        objectName: fixture.name,
        x: 0.5,
        y: 0.5 + index,
        w: 2,
        h: 1,
        shape: fixture.shape,
        rectRadius: fixture.radius,
      });
    }
    const imported = await openPptxGenJSPublicOutput(generated);
    const importedByName = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    for (const fixture of generatedCases) {
      const shape = importedByName.get(fixture.name)!;
      expect(directTextPresetGeometryState(shapeXml(imported, 0, shape.id)), fixture.name)
        .toEqual({
          type: fixture.expectedType,
          adjustments: fixture.expectedFormula === undefined
            ? []
            : [{ name: 'adj', formula: fixture.expectedFormula }],
        });
      if (fixture.expectedFormula?.includes('Infinity')) {
        expect(shape.adjustments, fixture.name).toBeUndefined();
      }
    }

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const existing = nativeSlide.addText('Existing', { name: 'Existing' });
    const explicitZero = nativeSlide.addText('Explicit zero radius', {
      shape: 'roundRect',
      rectRadius: inches(0),
      width: inches(2),
      height: inches(1),
    });
    expect(explicitZero.adjustments).toEqual([{ name: 'adj', value: 0 }]);
    const beforeInvalid = packageState(native);
    const beforeShapes = nativeSlide.shapes;
    const invalid = [
      { rectRadius: inches(0.5) },
      { shape: 'rect', rectRadius: inches(0.5) },
      { shape: 'ellipse', rectRadius: inches(0.5) },
      { shape: 'roundRect', rectRadius: '0.5' },
      { shape: 'roundRect', rectRadius: -1 },
      { shape: 'roundRect', rectRadius: inches(1) + 1 },
      { shape: 'roundRect', rectRadius: Number.NaN },
      { shape: 'roundRect', rectRadius: Number.POSITIVE_INFINITY },
      { shape: 'roundRect', rectRadius: Number.NEGATIVE_INFINITY },
    ];
    for (const options of invalid) {
      expect(() => nativeSlide.addText('Invalid native radius', options as never)).toThrow();
      expect(packageState(native)).toEqual(beforeInvalid);
      expect(nativeSlide.shapes).toEqual(beforeShapes);
      expect(nativeSlide.shapes[0]).toBe(existing);
      expect(nativeSlide.shapes[1]).toBe(explicitZero);
    }
  });

  it('compares shape fill public output and strict native divergences', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    try {
      const generatedCases: readonly {
        readonly name: string;
        readonly fill?: Record<string, unknown>;
      }[] = [
        { name: 'Fill omitted' },
        { name: 'Fill none', fill: { type: 'none' } },
        { name: 'Fill sRGB', fill: { color: 'FF0000' } },
        { name: 'Fill scheme', fill: { color: generated.SchemeColor.accent2 } },
        { name: 'Fill transparency', fill: { color: '00FF00', transparency: 50 } },
        { name: 'Fill zero', fill: { color: '0000FF', transparency: 0 } },
        { name: 'Fill deprecated alpha', fill: { color: '112233', alpha: 40 } },
        { name: 'Fill empty', fill: {} },
        { name: 'Fill missing color', fill: { type: 'solid' } },
      ];
      for (const fixture of generatedCases) {
        const options: Record<string, unknown> = { objectName: fixture.name };
        if (fixture.fill !== undefined) options.fill = fixture.fill;
        generatedSlide.addShape(generated.ShapeType.rect!, options);
      }

      const imported = await openPptxGenJSPublicOutput(generated);
      const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
        shape.name,
        shape as ShapeModel,
      ]));
      expect(importedShapes.size).toBe(generatedCases.length);

      const native = PptxDocument.create();
      const nativeSlide = native.addSlide();
      const nativeShapes = new Map([
        ['Fill omitted', nativeSlide.addShape('rect', { name: 'Fill omitted' })],
        ['Fill none', nativeSlide.addShape('rect', {
          name: 'Fill none',
          fill: { kind: 'none' },
        })],
        ['Fill sRGB', nativeSlide.addShape('rect', {
          name: 'Fill sRGB',
          fill: { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } },
        })],
        ['Fill scheme', nativeSlide.addShape('rect', {
          name: 'Fill scheme',
          fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' } },
        })],
        ['Fill transparency', nativeSlide.addShape('rect', {
          name: 'Fill transparency',
          fill: {
            kind: 'solid',
            color: { kind: 'srgb', value: '00FF00' },
            transparency: 50,
          },
        })],
        ['Fill zero', nativeSlide.addShape('rect', {
          name: 'Fill zero',
          fill: {
            kind: 'solid',
            color: { kind: 'srgb', value: '0000FF' },
            transparency: 0,
          },
        })],
      ] as const);

      for (const name of [
        'Fill omitted',
        'Fill sRGB',
        'Fill scheme',
        'Fill transparency',
      ] as const) {
        const importedShape = importedShapes.get(name)!;
        const nativeShape = nativeShapes.get(name)!;
        expect(importedShape).toBeInstanceOf(ShapeModel);
        expect(importedShape.name).toBe(nativeShape.name);
        expect(importedShape.presetType).toBe(nativeShape.presetType);
        expect(importedShape.transform).toEqual(nativeShape.transform);
        expect(importedShape.fill).toEqual(nativeShape.fill);
      }

      const generatedNone = importedShapes.get('Fill none')!;
      const nativeNone = nativeShapes.get('Fill none')!;
      expect(generatedNone.fill).toBeUndefined();
      expect(nativeNone.fill).toEqual({ kind: 'none' });
      expect(shapeXml(imported, 0, generatedNone.id)).toMatch(
        /<\/a:prstGeom><a:ln(?:\/>|>)/,
      );
      expect(shapeXml(native, 0, nativeNone.id)).toContain(
        '</a:prstGeom><a:noFill/><a:ln/>',
      );

      const generatedZero = importedShapes.get('Fill zero')!;
      const nativeZero = nativeShapes.get('Fill zero')!;
      expect(generatedZero.fill).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: '0000FF' },
      });
      expect(nativeZero.fill).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: '0000FF' },
        transparency: 0,
      });
      expect(shapeXml(imported, 0, generatedZero.id)).not.toContain('<a:alpha');
      expect(shapeXml(native, 0, nativeZero.id)).toContain('<a:alpha val="100000"/>');
      expect(100 - (generatedZero.fill?.kind === 'solid'
        ? generatedZero.fill.transparency ?? 0
        : 100)).toBe(100);
      expect(100 - (nativeZero.fill?.kind === 'solid'
        ? nativeZero.fill.transparency ?? 0
        : 100)).toBe(100);

      expect(importedShapes.get('Fill empty')!.fill).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: '000000' },
      });
      expect(importedShapes.get('Fill missing color')!.fill).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: '000000' },
      });
      expect(importedShapes.get('Fill deprecated alpha')!.fill).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: '112233' },
        transparency: 40,
      });
      expect(warning).toHaveBeenCalled();

      const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
      const invalidJournal = [...native.opcPackage.mutations];
      for (const fill of [
        {},
        { kind: 'solid' },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: '112233' },
          alpha: 40,
        },
      ]) {
        expect(() => nativeSlide.addShape('rect', { fill } as never)).toThrow();
      }
      expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
      expect(native.opcPackage.mutations).toEqual(invalidJournal);

      const editable = importedShapes.get('Fill sRGB')!;
      editable.fill = {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent4' },
        transparency: 25,
      };
      const reopened = await PptxDocument.open(await imported.write());
      const reopenedEditable = reopened.slides[0]!.shapes.find(
        ({ name }) => name === 'Fill sRGB',
      ) as ShapeModel;
      expect(reopenedEditable.fill).toEqual({
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent4' },
        transparency: 25,
      });
    } finally {
      warning.mockRestore();
    }
  });

  it('compares text shape fill public output and strict native divergences', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const generatedCases: readonly {
      readonly name: string;
      readonly fill?: Record<string, unknown>;
    }[] = [
      { name: 'Text fill omitted' },
      { name: 'Text fill none', fill: { type: 'none' } },
      {
        name: 'Text fill sRGB',
        fill: { color: 'AB12CD', transparency: 25 },
      },
      {
        name: 'Text fill scheme',
        fill: { color: generated.SchemeColor.accent2 },
      },
      {
        name: 'Text fill zero',
        fill: { color: '00AA00', transparency: 0 },
      },
      { name: 'Text fill missing color', fill: { type: 'solid' } },
    ];
    for (const [index, fixture] of generatedCases.entries()) {
      const options: Record<string, unknown> = {
        objectName: fixture.name,
        x: 1,
        y: 0.5 + index * 0.7,
        w: 4,
        h: 0.5,
      };
      if (fixture.fill !== undefined) options.fill = fixture.fill;
      generatedSlide.addText(fixture.name, options);
    }

    const warning = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    try {
      const imported = await openPptxGenJSPublicOutput(generated);
      const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
        shape.name,
        shape as ShapeModel,
      ]));
      expect([...importedShapes.keys()]).toEqual(generatedCases.map(({ name }) => name));
      expect(generatedCases.map(({ name }) => importedShapes.get(name)!.fill)).toEqual([
        { kind: 'none' },
        undefined,
        {
          kind: 'solid',
          color: { kind: 'srgb', value: 'AB12CD' },
          transparency: 25,
        },
        { kind: 'solid', color: { kind: 'scheme', value: 'accent2' } },
        { kind: 'solid', color: { kind: 'srgb', value: '00AA00' } },
        { kind: 'solid', color: { kind: 'srgb', value: '000000' } },
      ]);
      expect(warning).toHaveBeenCalled();

      const native = PptxDocument.create();
      const nativeSlide = native.addSlide();
      const nativeShapes = new Map([
        ['Text fill omitted', nativeSlide.addText('Text fill omitted', {
          name: 'Text fill omitted',
          x: inches(1),
          y: inches(0.5),
          width: inches(4),
          height: inches(0.5),
        })],
        ['Text fill none', nativeSlide.addText('Text fill none', {
          name: 'Text fill none',
          x: inches(1),
          y: inches(1.2),
          width: inches(4),
          height: inches(0.5),
          fill: { kind: 'none' },
        })],
        ['Text fill sRGB', nativeSlide.addText('Text fill sRGB', {
          name: 'Text fill sRGB',
          x: inches(1),
          y: inches(1.9),
          width: inches(4),
          height: inches(0.5),
          fill: {
            kind: 'solid',
            color: { kind: 'srgb', value: 'AB12CD' },
            transparency: 25,
          },
        })],
        ['Text fill scheme', nativeSlide.addText('Text fill scheme', {
          name: 'Text fill scheme',
          x: inches(1),
          y: inches(2.6),
          width: inches(4),
          height: inches(0.5),
          fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' } },
        })],
        ['Text fill zero', nativeSlide.addText('Text fill zero', {
          name: 'Text fill zero',
          x: inches(1),
          y: inches(3.3),
          width: inches(4),
          height: inches(0.5),
          fill: {
            kind: 'solid',
            color: { kind: 'srgb', value: '00AA00' },
            transparency: 0,
          },
        })],
      ] as const);

      for (const name of [
        'Text fill omitted',
        'Text fill sRGB',
        'Text fill scheme',
      ] as const) {
        const importedShape = importedShapes.get(name)!;
        const nativeShape = nativeShapes.get(name)!;
        expect(importedShape.name).toBe(nativeShape.name);
        expect(importedShape.text).toBe(nativeShape.text);
        expect(importedShape.transform).toEqual(nativeShape.transform);
        expect(importedShape.presetType).toBe(nativeShape.presetType);
        expect(importedShape.fill).toEqual(nativeShape.fill);
        expect(importedShape.line).toBeUndefined();
        expect(nativeShape.line).toEqual({ kind: 'none' });
      }

      const directFillState = (xml: string) => {
        const properties = xml.match(/<p:spPr(?:\s[^>]*)?>([\s\S]*?)<\/p:spPr>/)?.[1];
        if (!properties) throw new Error('Text shape properties were not found');
        const geometry = properties.match(/<a:prstGeom\b[^>]*\bprst="([^"]+)"/)?.[1];
        const afterGeometry = properties.match(/<\/a:prstGeom>([\s\S]*)$/)?.[1];
        if (afterGeometry === undefined) throw new Error('Text shape geometry was not found');
        const lineOffset = afterGeometry.indexOf('<a:ln');
        const fillXml = lineOffset < 0 ? afterGeometry : afterGeometry.slice(0, lineOffset);
        const solid = fillXml.match(
          /<a:solidFill><a:(srgbClr|schemeClr)\b[^>]*\bval="([^"]+)"(?:><a:alpha\b[^>]*\bval="([0-9]+)"\/><\/a:\1>|\/>)<\/a:solidFill>/,
        );
        return {
          geometry,
          fill: fillXml.includes('<a:noFill/>')
            ? { kind: 'none' as const }
            : solid
              ? {
                  kind: 'solid' as const,
                  colorKind: solid[1],
                  color: solid[2],
                  alpha: solid[3] === undefined ? undefined : Number(solid[3]),
                }
              : undefined,
          line: lineOffset >= 0,
        };
      };

      for (const name of [
        'Text fill omitted',
        'Text fill sRGB',
        'Text fill scheme',
      ] as const) {
        const importedShape = importedShapes.get(name)!;
        const nativeShape = nativeShapes.get(name)!;
        expect(directFillState(shapeXml(imported, 0, importedShape.id)))
          .toEqual(directFillState(shapeXml(native, 0, nativeShape.id)));
      }

      const generatedNone = importedShapes.get('Text fill none')!;
      const nativeNone = nativeShapes.get('Text fill none')!;
      expect(generatedNone.fill).toBeUndefined();
      expect(nativeNone.fill).toEqual({ kind: 'none' });
      expect(directFillState(shapeXml(imported, 0, generatedNone.id))).toEqual({
        geometry: 'rect',
        fill: undefined,
        line: true,
      });
      expect(directFillState(shapeXml(native, 0, nativeNone.id))).toEqual({
        geometry: 'rect',
        fill: { kind: 'none' },
        line: true,
      });

      const generatedZero = importedShapes.get('Text fill zero')!;
      const nativeZero = nativeShapes.get('Text fill zero')!;
      expect(generatedZero.fill).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: '00AA00' },
      });
      expect(nativeZero.fill).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: '00AA00' },
        transparency: 0,
      });
      expect(directFillState(shapeXml(imported, 0, generatedZero.id)).fill).toEqual({
        kind: 'solid',
        colorKind: 'srgbClr',
        color: '00AA00',
        alpha: undefined,
      });
      expect(directFillState(shapeXml(native, 0, nativeZero.id)).fill).toEqual({
        kind: 'solid',
        colorKind: 'srgbClr',
        color: '00AA00',
        alpha: 100000,
      });

      const generatedSrgb = importedShapes.get('Text fill sRGB')!;
      expect(directFillState(shapeXml(imported, 0, generatedSrgb.id)).fill).toEqual({
        kind: 'solid',
        colorKind: 'srgbClr',
        color: 'AB12CD',
        alpha: 75000,
      });
      const generatedScheme = importedShapes.get('Text fill scheme')!;
      expect(directFillState(shapeXml(imported, 0, generatedScheme.id)).fill).toEqual({
        kind: 'solid',
        colorKind: 'schemeClr',
        color: 'accent2',
        alpha: undefined,
      });

      const beforeInvalid = packageState(native);
      for (const fill of [
        { color: 'AB12CD' },
        { type: 'none' },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: 'AB12CD' },
          alpha: 25,
        },
        { kind: 'solid' },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: 'AB12CD' },
          transparency: '25',
        },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: 'AB12CD' },
          transparency: -1,
        },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: 'AB12CD' },
          transparency: 101,
        },
      ]) {
        expect(() => nativeSlide.addText('Invalid native fill', { fill } as never)).toThrow();
        expect(packageState(native)).toEqual(beforeInvalid);
      }

      const importedSlideBytes = imported.opcPackage
        .requirePart(imported.slides[0]!.partUri).bytes.slice();
      const reopened = await PptxDocument.open(await imported.write());
      expect(reopened.opcPackage.requirePart(reopened.slides[0]!.partUri).bytes)
        .toEqual(importedSlideBytes);
      expect(reopened.slides[0]!.shapes.map((shape) => (shape as ShapeModel).fill))
        .toEqual(imported.slides[0]!.shapes.map((shape) => (shape as ShapeModel).fill));
    } finally {
      warning.mockRestore();
    }
  });

  it('locks fill boundaries and deprecated alpha across shape and text owners', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const hosts: readonly {
      readonly name: string;
      readonly add: (name: string, fill: Record<string, unknown>) => void;
    }[] = [
      {
        name: 'shape rect',
        add: (name, fill) => generatedSlide.addShape(generated.ShapeType.rect!, {
          objectName: name,
          fill,
        }),
      },
      {
        name: 'shape line',
        add: (name, fill) => generatedSlide.addShape(generated.ShapeType.line!, {
          objectName: name,
          fill,
        }),
      },
      {
        name: 'text default',
        add: (name, fill) => generatedSlide.addText(name, { objectName: name, fill }),
      },
      {
        name: 'text ellipse',
        add: (name, fill) => generatedSlide.addText(name, {
          objectName: name,
          shape: generated.ShapeType.ellipse,
          fill,
        }),
      },
      {
        name: 'text line',
        add: (name, fill) => generatedSlide.addText(name, {
          objectName: name,
          shape: generated.ShapeType.line,
          fill,
          line: { color: '000000' },
        }),
      },
    ];
    const scenarios: readonly {
      readonly name: string;
      readonly fill: Record<string, unknown>;
      readonly expected?: ShapeFill;
      readonly alphaValues: readonly string[];
    }[] = [
      {
        name: 'explicit none',
        fill: { type: 'none' },
        alphaValues: [],
      },
      {
        name: 'explicit solid',
        fill: { type: 'solid', color: '001122' },
        expected: {
          kind: 'solid',
          color: { kind: 'srgb', value: '001122' },
        },
        alphaValues: [],
      },
      {
        name: 'scheme color',
        fill: { color: generated.SchemeColor.accent2 },
        expected: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
        },
        alphaValues: [],
      },
      {
        name: 'full transparency',
        fill: { color: '102030', transparency: 100 },
        expected: {
          kind: 'solid',
          color: { kind: 'srgb', value: '102030' },
          transparency: 100,
        },
        alphaValues: ['0'],
      },
      {
        name: 'fractional transparency',
        fill: { color: '203040', transparency: 33.333 },
        expected: {
          kind: 'solid',
          color: { kind: 'srgb', value: '203040' },
          transparency: 33.333,
        },
        alphaValues: ['66667'],
      },
      {
        name: 'alpha only',
        fill: { color: '112233', alpha: 40 },
        expected: {
          kind: 'solid',
          color: { kind: 'srgb', value: '112233' },
          transparency: 40,
        },
        alphaValues: ['60000'],
      },
      {
        name: 'transparency then alpha',
        fill: { color: '223344', transparency: 25, alpha: 40 },
        alphaValues: ['60000', '75000'],
      },
      {
        name: 'zero alpha',
        fill: { color: '334455', transparency: 25, alpha: 0 },
        expected: {
          kind: 'solid',
          color: { kind: 'srgb', value: '334455' },
          transparency: 25,
        },
        alphaValues: ['75000'],
      },
      {
        name: 'zero transparency',
        fill: { color: '445566', transparency: 0, alpha: 40 },
        expected: {
          kind: 'solid',
          color: { kind: 'srgb', value: '445566' },
          transparency: 40,
        },
        alphaValues: ['60000'],
      },
    ];
    const expected = new Map<string, (typeof scenarios)[number]>();
    for (const host of hosts) {
      for (const scenario of scenarios) {
        const name = `${host.name} ${scenario.name}`;
        host.add(name, scenario.fill);
        expected.set(name, scenario);
      }
    }

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    expect(importedShapes.size).toBe(expected.size);
    for (const [name, scenario] of expected) {
      const shape = importedShapes.get(name)!;
      expect(shape, name).toBeInstanceOf(ShapeModel);
      expect(shape.fill, name).toEqual(scenario.expected);
      const alphaValues = [...shapeXml(imported, 0, shape.id).matchAll(
        /<a:alpha val="([0-9]+)"\/>/g,
      )].map((match) => match[1]);
      expect(alphaValues, name).toEqual(scenario.alphaValues);
    }
  });

  it('compares text shape line public output and strict native divergences', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const dashes = [
      'solid',
      'dash',
      'dashDot',
      'lgDash',
      'lgDashDot',
      'lgDashDotDot',
      'sysDash',
      'sysDot',
    ] as const;
    const generatedCases: readonly {
      readonly name: string;
      readonly line?: Record<string, unknown>;
    }[] = [
      { name: 'Text line omitted' },
      { name: 'Text line none', line: { type: 'none' } },
      { name: 'Text line empty', line: {} },
      { name: 'Text line missing color', line: { type: 'solid' } },
      { name: 'Text line sRGB', line: { color: 'AB12CD' } },
      { name: 'Text line scheme', line: { color: generated.SchemeColor.accent2 } },
      {
        name: 'Text line transparency',
        line: { color: '00FF00', transparency: 25 },
      },
      {
        name: 'Text line zero transparency',
        line: { color: '0000FF', transparency: 0 },
      },
      {
        name: 'Text line full transparency',
        line: { color: '00AA00', transparency: 100 },
      },
      { name: 'Text line zero width', line: { color: '112233', width: 0 } },
      { name: 'Text line positive width', line: { color: '223344', width: 2.5 } },
      { name: 'Text line deprecated alpha', line: { color: '334455', alpha: 40 } },
      { name: 'Text line deprecated dash', line: { color: '445566', lineDash: 'dash' } },
      ...dashes.map((dash) => ({
        name: `Text line dash ${dash}`,
        line: { color: '556677', dashType: dash },
      })),
    ];
    for (const [index, fixture] of generatedCases.entries()) {
      const options: Record<string, unknown> = {
        objectName: fixture.name,
        x: 1,
        y: 0.25 + index * 0.25,
        w: 4,
        h: 0.2,
      };
      if (fixture.line !== undefined) options.line = fixture.line;
      generatedSlide.addText(fixture.name, options);
    }

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    expect([...importedShapes.keys()]).toEqual(generatedCases.map(({ name }) => name));
    expect(generatedCases.map(({ name }) => importedShapes.get(name)!.line)).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
      {
        kind: 'line',
        color: { kind: 'srgb', value: 'AB12CD' },
        width: 1,
        dash: 'solid',
      },
      {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent2' },
        width: 1,
        dash: 'solid',
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: '00FF00' },
        transparency: 25,
        width: 1,
        dash: 'solid',
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: '0000FF' },
        width: 1,
        dash: 'solid',
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: '00AA00' },
        transparency: 100,
        width: 1,
        dash: 'solid',
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
        width: 1,
        dash: 'solid',
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: '223344' },
        width: 2.5,
        dash: 'solid',
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: '334455' },
        transparency: 40,
        width: 1,
        dash: 'solid',
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: '445566' },
        width: 1,
        dash: 'solid',
      },
      ...dashes.map((dash) => ({
        kind: 'line' as const,
        color: { kind: 'srgb' as const, value: '556677' },
        width: 1,
        dash,
      })),
    ]);

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const nativeShapes = new Map<string, ShapeModel>([
      ['Text line omitted', nativeSlide.addText('Text line omitted', {
        name: 'Text line omitted',
        x: inches(1),
        y: inches(0.25),
        width: inches(4),
        height: inches(0.2),
      })],
      ['Text line none', nativeSlide.addText('Text line none', {
        name: 'Text line none',
        x: inches(1),
        y: inches(0.5),
        width: inches(4),
        height: inches(0.2),
        line: { kind: 'none' },
      })],
      ['Text line sRGB', nativeSlide.addText('Text line sRGB', {
        name: 'Text line sRGB',
        x: inches(1),
        y: inches(1.25),
        width: inches(4),
        height: inches(0.2),
        line: { kind: 'line', color: { kind: 'srgb', value: 'AB12CD' } },
      })],
      ['Text line scheme', nativeSlide.addText('Text line scheme', {
        name: 'Text line scheme',
        x: inches(1),
        y: inches(1.5),
        width: inches(4),
        height: inches(0.2),
        line: { kind: 'line', color: { kind: 'scheme', value: 'accent2' } },
      })],
      ['Text line transparency', nativeSlide.addText('Text line transparency', {
        name: 'Text line transparency',
        x: inches(1),
        y: inches(1.75),
        width: inches(4),
        height: inches(0.2),
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '00FF00' },
          transparency: 25,
        },
      })],
      ['Text line zero transparency', nativeSlide.addText('Text line zero transparency', {
        name: 'Text line zero transparency',
        x: inches(1),
        y: inches(2),
        width: inches(4),
        height: inches(0.2),
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '0000FF' },
          transparency: 0,
        },
      })],
      ['Text line full transparency', nativeSlide.addText('Text line full transparency', {
        name: 'Text line full transparency',
        x: inches(1),
        y: inches(2.25),
        width: inches(4),
        height: inches(0.2),
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '00AA00' },
          transparency: 100,
        },
      })],
      ['Text line zero width', nativeSlide.addText('Text line zero width', {
        name: 'Text line zero width',
        x: inches(1),
        y: inches(2.5),
        width: inches(4),
        height: inches(0.2),
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '112233' },
          width: 0,
        },
      })],
      ['Text line positive width', nativeSlide.addText('Text line positive width', {
        name: 'Text line positive width',
        x: inches(1),
        y: inches(2.75),
        width: inches(4),
        height: inches(0.2),
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '223344' },
          width: 2.5,
        },
      })],
      ...dashes.map((dash, dashIndex) => [
        `Text line dash ${dash}`,
        nativeSlide.addText(`Text line dash ${dash}`, {
          name: `Text line dash ${dash}`,
          x: inches(1),
          y: inches(3.5 + dashIndex * 0.25),
          width: inches(4),
          height: inches(0.2),
          line: {
            kind: 'line',
            color: { kind: 'srgb', value: '556677' },
            dash,
          },
        }),
      ] as const),
    ] as const);

    for (const name of [
      'Text line sRGB',
      'Text line scheme',
      'Text line transparency',
      'Text line full transparency',
      'Text line positive width',
      ...dashes.map((dash) => `Text line dash ${dash}`),
    ]) {
      const importedShape = importedShapes.get(name)!;
      const nativeShape = nativeShapes.get(name)!;
      expect(importedShape.name, name).toBe(nativeShape.name);
      expect(importedShape.text, name).toBe(nativeShape.text);
      expect(importedShape.transform, name).toEqual(nativeShape.transform);
      expect(importedShape.presetType, name).toBe(nativeShape.presetType);
      expect(importedShape.fill, name).toEqual(nativeShape.fill);
      expect(importedShape.line, name).toEqual(nativeShape.line);
    }

    const directLineState = (xml: string) => {
      const properties = xml.match(/<p:spPr(?:\s[^>]*)?>([\s\S]*?)<\/p:spPr>/)?.[1];
      if (!properties) throw new Error('Text shape properties were not found');
      const fillOffset = properties.search(/<a:(?:noFill|solidFill)\b/);
      const lineOffset = properties.indexOf('<a:ln');
      if (fillOffset < 0 || lineOffset < 0 || fillOffset >= lineOffset) {
        throw new Error('Text shape fill and line order is invalid');
      }
      const lineXml = properties.slice(lineOffset);
      const width = lineXml.match(/^<a:ln\b[^>]*\bw="([0-9]+)"/)?.[1];
      const solid = lineXml.match(
        /<a:solidFill><a:(srgbClr|schemeClr)\b[^>]*\bval="([^"]+)"(?:><a:alpha\b[^>]*\bval="([0-9]+)"\/><\/a:\1>|\/>)<\/a:solidFill>/,
      );
      return {
        width: width === undefined ? undefined : Number(width),
        noFill: lineXml.includes('<a:noFill/>'),
        colorKind: solid?.[1],
        color: solid?.[2],
        alpha: solid?.[3] === undefined ? undefined : Number(solid[3]),
        dash: lineXml.match(/<a:prstDash\b[^>]*\bval="([^"]+)"\/>/)?.[1],
      };
    };

    for (const name of ['Text line omitted', 'Text line none'] as const) {
      const importedShape = importedShapes.get(name)!;
      const nativeShape = nativeShapes.get(name)!;
      expect(importedShape.line).toBeUndefined();
      expect(nativeShape.line).toEqual({ kind: 'none' });
      expect(directLineState(shapeXml(imported, 0, importedShape.id))).toEqual({
        width: undefined,
        noFill: false,
        colorKind: undefined,
        color: undefined,
        alpha: undefined,
        dash: undefined,
      });
      expect(directLineState(shapeXml(native, 0, nativeShape.id))).toEqual({
        width: undefined,
        noFill: true,
        colorKind: undefined,
        color: undefined,
        alpha: undefined,
        dash: undefined,
      });
    }
    for (const name of ['Text line empty', 'Text line missing color']) {
      expect(importedShapes.get(name)!.line, name).toBeUndefined();
    }

    const generatedSrgb = importedShapes.get('Text line sRGB')!;
    const nativeSrgb = nativeShapes.get('Text line sRGB')!;
    expect(directLineState(shapeXml(imported, 0, generatedSrgb.id))).toEqual({
      width: undefined,
      noFill: false,
      colorKind: 'srgbClr',
      color: 'AB12CD',
      alpha: undefined,
      dash: undefined,
    });
    expect(directLineState(shapeXml(native, 0, nativeSrgb.id))).toEqual({
      width: 12_700,
      noFill: false,
      colorKind: 'srgbClr',
      color: 'AB12CD',
      alpha: undefined,
      dash: 'solid',
    });

    const generatedZeroTransparency = importedShapes.get('Text line zero transparency')!;
    const nativeZeroTransparency = nativeShapes.get('Text line zero transparency')!;
    expect(generatedZeroTransparency.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '0000FF' },
      width: 1,
      dash: 'solid',
    });
    expect(nativeZeroTransparency.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '0000FF' },
      transparency: 0,
      width: 1,
      dash: 'solid',
    });
    expect(directLineState(shapeXml(imported, 0, generatedZeroTransparency.id)).alpha)
      .toBeUndefined();
    expect(directLineState(shapeXml(native, 0, nativeZeroTransparency.id)).alpha)
      .toBe(100_000);

    const generatedZeroWidth = importedShapes.get('Text line zero width')!;
    const nativeZeroWidth = nativeShapes.get('Text line zero width')!;
    expect(generatedZeroWidth.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      width: 1,
      dash: 'solid',
    });
    expect(nativeZeroWidth.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      width: 0,
      dash: 'solid',
    });
    expect(directLineState(shapeXml(imported, 0, generatedZeroWidth.id)).width)
      .toBeUndefined();
    expect(directLineState(shapeXml(native, 0, nativeZeroWidth.id)).width).toBe(0);

    expect(importedShapes.get('Text line deprecated alpha')!.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '334455' },
      transparency: 40,
      width: 1,
      dash: 'solid',
    });
    expect(importedShapes.get('Text line deprecated dash')!.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '445566' },
      width: 1,
      dash: 'solid',
    });

    const beforeInvalid = packageState(native);
    for (const line of [
      {},
      { type: 'none' },
      { color: 'AB12CD' },
      { kind: 'line' },
      {
        kind: 'line',
        color: { kind: 'srgb', value: 'AB12CD' },
        alpha: 25,
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: 'AB12CD' },
        dashType: 'dash',
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: 'AB12CD' },
        transparency: '25',
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: 'AB12CD' },
        width: 1_585,
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: 'AB12CD' },
        dash: 'dot',
      },
    ]) {
      expect(() => nativeSlide.addText('Invalid native line', { line } as never)).toThrow();
      expect(packageState(native)).toEqual(beforeInvalid);
    }

    const importedSlideBytes = imported.opcPackage
      .requirePart(imported.slides[0]!.partUri).bytes.slice();
    const reopened = await PptxDocument.open(await imported.write());
    expect(reopened.opcPackage.requirePart(reopened.slides[0]!.partUri).bytes)
      .toEqual(importedSlideBytes);
    expect(reopened.slides[0]!.shapes.map((shape) => (shape as ShapeModel).line))
      .toEqual(imported.slides[0]!.shapes.map((shape) => (shape as ShapeModel).line));
  });

  it('compares text shape arrows public output and strict native divergences', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const arrowTypes = [
      'none',
      'arrow',
      'diamond',
      'oval',
      'stealth',
      'triangle',
    ] as const;
    const supportedCases: {
      readonly name: string;
      readonly generatedLine: Record<string, unknown>;
      readonly line: ShapeLine;
      readonly arrows: ShapeArrows;
    }[] = [];
    for (const type of arrowTypes) {
      supportedCases.push(
        {
          name: `Text arrow begin ${type}`,
          generatedLine: { color: '112233', beginArrowType: type },
          line: { kind: 'line', color: { kind: 'srgb', value: '112233' } },
          arrows: { begin: type },
        },
        {
          name: `Text arrow end ${type}`,
          generatedLine: { color: '112233', endArrowType: type },
          line: { kind: 'line', color: { kind: 'srgb', value: '112233' } },
          arrows: { end: type },
        },
      );
    }
    supportedCases.push({
      name: 'Text arrow solid both',
      generatedLine: {
        color: '112233',
        width: 2.5,
        dashType: 'dashDot',
        beginArrowType: 'stealth',
        endArrowType: 'oval',
      },
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
        width: 2.5,
        dash: 'dashDot',
      },
      arrows: { begin: 'stealth', end: 'oval' },
    });
    const generatedCases: readonly {
      readonly name: string;
      readonly line?: Record<string, unknown>;
      readonly options?: Record<string, unknown>;
    }[] = [
      { name: 'Text arrow omitted' },
      ...supportedCases.map(({ name, generatedLine }) => ({ name, line: generatedLine })),
      {
        name: 'Text arrow only',
        line: { beginArrowType: 'triangle', endArrowType: 'arrow' },
      },
      {
        name: 'Text arrow no line',
        line: { type: 'none', beginArrowType: 'diamond' },
      },
      {
        name: 'Text arrow empty ignored',
        line: { color: '112233', beginArrowType: '', endArrowType: '' },
      },
      {
        name: 'Text arrow nested aliases ignored',
        line: { color: '112233', lineHead: 'triangle', lineTail: 'arrow' },
      },
      {
        name: 'Text arrow top aliases ignored',
        line: { color: '112233' },
        options: { lineHead: 'triangle', lineTail: 'arrow' },
      },
      {
        name: 'Text arrow invalid passthrough',
        line: { color: '112233', beginArrowType: 'bogus' },
      },
    ];
    for (const [index, fixture] of generatedCases.entries()) {
      const options: Record<string, unknown> = {
        objectName: fixture.name,
        x: 1,
        y: 0.2 + index * 0.25,
        w: 4,
        h: 0.2,
        ...fixture.options,
      };
      if (fixture.line !== undefined) options.line = fixture.line;
      generatedSlide.addText(fixture.name, options);
    }

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    expect([...importedShapes.keys()]).toEqual(generatedCases.map(({ name }) => name));

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const nativeShapes = new Map<string, ShapeModel>();
    const addNative = (
      name: string,
      options: { readonly line?: ShapeLine; readonly arrows?: ShapeArrows },
    ) => {
      const index = generatedCases.findIndex((fixture) => fixture.name === name);
      if (index < 0) throw new Error(`Generated text arrow case ${name} was not found`);
      const shape = nativeSlide.addText(name, {
        name,
        x: inches(1),
        y: inches(0.2 + index * 0.25),
        width: inches(4),
        height: inches(0.2),
        ...options,
      });
      nativeShapes.set(name, shape);
      return shape;
    };
    addNative('Text arrow omitted', {});
    for (const fixture of supportedCases) {
      addNative(fixture.name, { line: fixture.line, arrows: fixture.arrows });
    }
    addNative('Text arrow only', {
      arrows: { begin: 'triangle', end: 'arrow' },
    });
    addNative('Text arrow no line', {
      line: { kind: 'none' },
      arrows: { begin: 'diamond' },
    });

    for (const fixture of supportedCases) {
      const name = fixture.name;
      const importedShape = importedShapes.get(name)!;
      const nativeShape = nativeShapes.get(name)!;
      expect(importedShape.name, name).toBe(nativeShape.name);
      expect(importedShape.text, name).toBe(nativeShape.text);
      expect(importedShape.transform, name).toEqual(nativeShape.transform);
      expect(importedShape.presetType, name).toBe(nativeShape.presetType);
      expect(importedShape.line, name).toEqual(nativeShape.line);
      expect(importedShape.arrows, name).toEqual(nativeShape.arrows);
      const importedXml = shapeXml(imported, 0, importedShape.id);
      const nativeXml = shapeXml(native, 0, nativeShape.id);
      if (fixture.arrows.begin !== undefined) {
        const endpoint = `<a:headEnd type=\"${fixture.arrows.begin}\"/>`;
        expect(importedXml, name).toContain(endpoint);
        expect(nativeXml, name).toContain(endpoint);
      }
      if (fixture.arrows.end !== undefined) {
        const endpoint = `<a:tailEnd type=\"${fixture.arrows.end}\"/>`;
        expect(importedXml, name).toContain(endpoint);
        expect(nativeXml, name).toContain(endpoint);
      }
    }

    const solidImported = importedShapes.get('Text arrow solid both')!;
    const solidNative = nativeShapes.get('Text arrow solid both')!;
    for (const xml of [
      shapeXml(imported, 0, solidImported.id),
      shapeXml(native, 0, solidNative.id),
    ]) {
      expect(xml).toContain('<a:ln w="31750">');
      expect(xml).toContain('<a:prstDash val="dashDot"/>');
      const fillOffset = xml.indexOf('<a:solidFill>');
      const dashOffset = xml.indexOf('<a:prstDash');
      const headOffset = xml.indexOf('<a:headEnd');
      const tailOffset = xml.indexOf('<a:tailEnd');
      expect(fillOffset).toBeGreaterThanOrEqual(0);
      expect(fillOffset).toBeLessThan(dashOffset);
      expect(dashOffset).toBeLessThan(headOffset);
      expect(headOffset).toBeLessThan(tailOffset);
    }

    const importedOmitted = importedShapes.get('Text arrow omitted')!;
    const nativeOmitted = nativeShapes.get('Text arrow omitted')!;
    expect(importedOmitted.arrows).toBeUndefined();
    expect(nativeOmitted.arrows).toBeUndefined();
    expect(importedOmitted.line).toBeUndefined();
    expect(nativeOmitted.line).toEqual({ kind: 'none' });
    expect(shapeXml(imported, 0, importedOmitted.id)).toContain('<a:ln></a:ln>');
    expect(shapeXml(native, 0, nativeOmitted.id)).toContain('<a:ln><a:noFill/></a:ln>');

    const importedArrowOnly = importedShapes.get('Text arrow only')!;
    const nativeArrowOnly = nativeShapes.get('Text arrow only')!;
    expect(importedArrowOnly.arrows).toEqual({ begin: 'triangle', end: 'arrow' });
    expect(nativeArrowOnly.arrows).toEqual(importedArrowOnly.arrows);
    expect(importedArrowOnly.line).toBeUndefined();
    expect(nativeArrowOnly.line).toEqual({ kind: 'none' });
    expect(shapeXml(imported, 0, importedArrowOnly.id))
      .toContain('<a:ln><a:headEnd type="triangle"/><a:tailEnd type="arrow"/></a:ln>');
    expect(shapeXml(native, 0, nativeArrowOnly.id))
      .toContain(
        '<a:ln><a:noFill/><a:headEnd type="triangle"/><a:tailEnd type="arrow"/></a:ln>',
      );

    const importedNoLine = importedShapes.get('Text arrow no line')!;
    const nativeNoLine = nativeShapes.get('Text arrow no line')!;
    expect(importedNoLine.arrows).toEqual({ begin: 'diamond' });
    expect(nativeNoLine.arrows).toEqual(importedNoLine.arrows);
    expect(importedNoLine.line).toBeUndefined();
    expect(nativeNoLine.line).toEqual({ kind: 'none' });
    expect(shapeXml(imported, 0, importedNoLine.id))
      .toContain('<a:ln><a:headEnd type="diamond"/></a:ln>');
    expect(shapeXml(native, 0, nativeNoLine.id))
      .toContain('<a:ln><a:noFill/><a:headEnd type="diamond"/></a:ln>');

    for (const name of [
      'Text arrow empty ignored',
      'Text arrow nested aliases ignored',
      'Text arrow top aliases ignored',
    ]) {
      const shape = importedShapes.get(name)!;
      expect(shape.arrows, name).toBeUndefined();
      expect(shapeXml(imported, 0, shape.id), name).not.toMatch(/<a:(?:headEnd|tailEnd)\b/);
    }
    const invalidImported = importedShapes.get('Text arrow invalid passthrough')!;
    expect(invalidImported.arrows).toBeUndefined();
    expect(shapeXml(imported, 0, invalidImported.id))
      .toContain('<a:headEnd type="bogus"/>');

    const beforeInvalid = packageState(native);
    for (const arrows of [
      { begin: '' },
      { end: 'bogus' },
      { beginArrowType: 'arrow' },
      { endArrowType: 'arrow' },
      { lineHead: 'arrow' },
      { lineTail: 'arrow' },
    ]) {
      expect(() => nativeSlide.addText('Invalid native text arrows', { arrows } as never))
        .toThrow(TypeError);
      expect(packageState(native)).toEqual(beforeInvalid);
    }

    const editable = importedShapes.get('Text arrow solid both')!;
    const importedLine = editable.line;
    editable.arrows = { begin: 'diamond' };
    expect(editable.line).toEqual(importedLine);
    editable.line = { kind: 'none' };
    expect(editable.arrows).toEqual({ begin: 'diamond' });
    editable.arrows = { end: 'stealth' };
    expect(editable.line).toEqual({ kind: 'none' });

    const invalidXmlBefore = shapeXml(imported, 0, invalidImported.id);
    const reopened = await PptxDocument.open(await imported.write());
    const reopenedEditable = reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'Text arrow solid both',
    ) as ShapeModel;
    expect(reopenedEditable.arrows).toEqual({ end: 'stealth' });
    expect(reopenedEditable.line).toEqual({ kind: 'none' });
    expect(shapeXml(reopened, 0, reopenedEditable.id))
      .toContain('<a:ln><a:noFill/><a:tailEnd type="stealth"/></a:ln>');
    const reopenedInvalid = reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'Text arrow invalid passthrough',
    ) as ShapeModel;
    expect(reopenedInvalid.arrows).toBeUndefined();
    expect(shapeXml(reopened, 0, reopenedInvalid.id)).toBe(invalidXmlBefore);
  });

  it('compares text shape shadow public output and strict native divergences', async () => {
    expect(new PptxGenJS().version).toBe('4.0.1');
    const generate = async (
      name: string,
      shadow: unknown,
      extra: Record<string, unknown> = {},
    ) => {
      const generated = new PptxGenJS();
      const slide = generated.addSlide();
      const options: Record<string, unknown> = {
        objectName: name,
        x: 1,
        y: 1,
        w: 2,
        h: 1,
        ...extra,
      };
      if (shadow !== undefined) options.shadow = shadow;
      slide.addText(name, options);
      const bytes = await generated.write({
        outputType: 'nodebuffer',
        compression: true,
      });
      const document = await PptxDocument.open(bytes);
      const partUri = document.slides[0]!.partUri;
      const xml = new TextDecoder().decode(document.opcPackage.requirePart(partUri).bytes);
      return { document, xml };
    };

    const omitted = await generate('Text shadow omitted', undefined);
    const none = await generate('Text shadow none', { type: 'none' });
    for (const fixture of [omitted, none]) {
      expect(fixture.xml).not.toContain('<a:effectLst>');
      expect(fixture.xml).not.toMatch(/<(?:a:)?(?:inner|outer)Shdw/);
      expect((fixture.document.slides[0]!.shapes[0] as ShapeModel).shadow).toBeUndefined();
    }

    const defaults = await generate('Text shadow outer defaults', { type: 'outer' });
    expect(defaults.xml).toContain(
      '<a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="bl" ' +
      'rotWithShape="0" blurRad="101600" dist="50800" dir="16200000"> ' +
      '<a:srgbClr val="000000"> <a:alpha val="75000"/></a:srgbClr> ' +
      '</a:outerShdw>',
    );
    const generatedDefaults = (defaults.document.slides[0]!.shapes[0] as ShapeModel).shadow;
    expect(generatedDefaults).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '000000' },
      opacity: 0.75,
      blur: 8,
      angle: 270,
      distance: 4,
      rotateWithShape: false,
    });

    const custom = await generate('Text shadow outer custom', {
      type: 'outer',
      color: '123ABC',
      opacity: 0.42,
      blur: 7.25,
      angle: 123,
      offset: 5.5,
      rotateWithShape: false,
    }, {
      line: {
        color: '112233',
        width: 2.5,
        dashType: 'dashDot',
        beginArrowType: 'triangle',
        endArrowType: 'arrow',
      },
    });
    const generatedCustomShape = custom.document.slides[0]!.shapes[0] as ShapeModel;
    expect(generatedCustomShape.shadow).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '123ABC' },
      opacity: 0.42,
      blur: 7.25,
      angle: 123,
      distance: 5.5,
      rotateWithShape: false,
    });
    expect(generatedCustomShape.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      width: 2.5,
      dash: 'dashDot',
    });
    expect(generatedCustomShape.arrows).toEqual({ begin: 'triangle', end: 'arrow' });
    const generatedLineOffset = custom.xml.indexOf('<a:ln w="31750">');
    const generatedEffectOffset = custom.xml.indexOf('<a:effectLst>', generatedLineOffset);
    expect(generatedLineOffset).toBeGreaterThanOrEqual(0);
    expect(generatedLineOffset).toBeLessThan(generatedEffectOffset);
    expect(custom.xml.slice(generatedLineOffset, generatedEffectOffset))
      .toContain('<a:headEnd type="triangle"/><a:tailEnd type="arrow"/>');

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const addNative = (name: string, shadow: ShapeShadow, options: {
      readonly line?: ShapeLine;
      readonly arrows?: ShapeArrows;
    } = {}) => nativeSlide.addText(name, {
      name,
      x: inches(1),
      y: inches(1),
      width: inches(2),
      height: inches(1),
      shadow,
      ...options,
    });
    const nativeDefaults = addNative('Text shadow outer defaults', { kind: 'outer' });
    const nativeCustom = addNative('Text shadow outer custom', {
      kind: 'outer',
      color: { kind: 'srgb', value: '123ABC' },
      opacity: 0.42,
      blur: 7.25,
      angle: 123,
      distance: 5.5,
      rotateWithShape: false,
    }, {
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
        width: 2.5,
        dash: 'dashDot',
      },
      arrows: { begin: 'triangle', end: 'arrow' },
    });
    expect(nativeDefaults.shadow).toEqual(generatedDefaults);
    expect(nativeCustom.name).toBe(generatedCustomShape.name);
    expect(nativeCustom.text).toBe(generatedCustomShape.text);
    expect(nativeCustom.transform).toEqual(generatedCustomShape.transform);
    expect(nativeCustom.presetType).toBe(generatedCustomShape.presetType);
    expect(nativeCustom.line).toEqual(generatedCustomShape.line);
    expect(nativeCustom.arrows).toEqual(generatedCustomShape.arrows);
    expect(nativeCustom.shadow).toEqual(generatedCustomShape.shadow);
    const nativeCustomXml = shapeXml(native, 0, nativeCustom.id);
    const nativeLineOffset = nativeCustomXml.indexOf('<a:ln w="31750">');
    const nativeEffectOffset = nativeCustomXml.indexOf('<a:effectLst>', nativeLineOffset);
    expect(nativeLineOffset).toBeGreaterThanOrEqual(0);
    expect(nativeLineOffset).toBeLessThan(nativeEffectOffset);

    const zero = await generate('Text shadow zero fallback', {
      type: 'outer',
      color: '000000',
      opacity: 0,
      blur: 0,
      angle: 0,
      offset: 0,
    });
    expect((zero.document.slides[0]!.shapes[0] as ShapeModel).shadow)
      .toEqual(generatedDefaults);
    const nativeZero = addNative('Text shadow native zero', {
      kind: 'outer',
      color: { kind: 'scheme', value: 'accent2' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
    });
    expect(nativeZero.shadow).toEqual({
      kind: 'outer',
      color: { kind: 'scheme', value: 'accent2' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
      rotateWithShape: false,
    });

    const rotate = await generate('Text shadow rotate ignored', {
      type: 'outer',
      rotateWithShape: true,
    });
    expect((rotate.document.slides[0]!.shapes[0] as ShapeModel).shadow)
      .toMatchObject({ kind: 'outer', rotateWithShape: false });
    const nativeRotate = addNative('Text shadow rotate kept', {
      kind: 'outer',
      rotateWithShape: true,
    });
    expect(nativeRotate.shadow).toMatchObject({ kind: 'outer', rotateWithShape: true });

    const malformedInner = await generate('Text shadow malformed inner', { type: 'inner' });
    expect(malformedInner.xml).toContain('<a:innerShdw  blurRad="101600"');
    expect(malformedInner.xml).toContain('</a:outerShdw></a:effectLst>');
    expect(() => malformedInner.document.slides[0]!.shapes).toThrow();
    await expect(malformedInner.document.write()).rejects.toThrow();
    const nativeInner = addNative('Text shadow native inner', {
      kind: 'inner',
      color: { kind: 'scheme', value: 'accent3' },
    });
    expect(nativeInner.shadow?.kind).toBe('inner');
    expect(shapeXml(native, 0, nativeInner.id)).toContain(
      '<a:innerShdw blurRad="101600" dist="50800" dir="16200000">',
    );
    expect(shapeXml(native, 0, nativeInner.id)).toContain('</a:innerShdw>');

    const warnings = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    const missingType = await generate('Text shadow missing type', {});
    const invalidType = await generate('Text shadow invalid type', {
      type: 'bogus',
      color: 'FF0000',
    });
    const hashColor = await generate('Text shadow hash color', {
      type: 'outer',
      color: '#ABCDEF',
    });
    const coercion = await generate('Text shadow coercion', {
      type: 'outer',
      color: '112233',
      opacity: '0.4',
      blur: '2.5',
      angle: '45.6',
      offset: '3',
    });
    const invalidRanges = await generate('Text shadow invalid ranges', {
      type: 'outer',
      color: '00FF00',
      opacity: 2,
      blur: -1,
      angle: 400,
      offset: 201,
    });
    expect(warnings).toHaveBeenCalled();
    warnings.mockRestore();

    expect((missingType.document.slides[0]!.shapes[0] as ShapeModel).shadow)
      .toEqual(generatedDefaults);
    expect((invalidType.document.slides[0]!.shapes[0] as ShapeModel).shadow)
      .toMatchObject({ kind: 'outer', color: { kind: 'srgb', value: 'FF0000' } });
    expect((hashColor.document.slides[0]!.shapes[0] as ShapeModel).shadow)
      .toMatchObject({ kind: 'outer', color: { kind: 'srgb', value: 'ABCDEF' } });
    const nativeHashColor = addNative('Text shadow native hash color', {
      kind: 'outer',
      color: { kind: 'srgb', value: '#ABCDEF' },
    });
    expect(nativeHashColor.shadow).toEqual(
      (hashColor.document.slides[0]!.shapes[0] as ShapeModel).shadow,
    );
    expect((coercion.document.slides[0]!.shapes[0] as ShapeModel).shadow).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '112233' },
      opacity: 0.4,
      blur: 2.5,
      angle: 46,
      distance: 3,
      rotateWithShape: false,
    });
    expect(invalidRanges.xml).toContain(
      'blurRad="-12700" dist="2552700" dir="16200000"> ' +
      '<a:srgbClr val="00FF00"> <a:alpha val="75000"/>',
    );
    expect((invalidRanges.document.slides[0]!.shapes[0] as ShapeModel).shadow)
      .toBeUndefined();

    const beforeInvalid = packageState(native);
    let accessorCalls = 0;
    const accessor = Object.defineProperty({}, 'kind', {
      enumerable: true,
      get() {
        accessorCalls += 1;
        return 'outer';
      },
    });
    for (const [index, shadow] of [
      {},
      { kind: 'none' },
      { type: 'outer' },
      { kind: 'outer', offset: 4 },
      { kind: 'inner', rotateWithShape: false },
      { kind: 'outer', opacity: '0.4' },
      { kind: 'outer', blur: -1 },
      { kind: 'outer', angle: 360 },
      { kind: 'outer', distance: 201 },
      accessor,
    ].entries()) {
      expect(
        () => nativeSlide.addText('Invalid native text shadow', { shadow } as never),
        `invalid native text shadow ${index}`,
      )
        .toThrow();
      expect(packageState(native)).toEqual(beforeInvalid);
    }
    expect(accessorCalls).toBe(0);

    const importedLine = generatedCustomShape.line;
    const importedArrows = generatedCustomShape.arrows;
    generatedCustomShape.shadow = undefined;
    expect(generatedCustomShape.line).toEqual(importedLine);
    expect(generatedCustomShape.arrows).toEqual(importedArrows);
    generatedCustomShape.shadow = { kind: 'outer', rotateWithShape: true };
    generatedCustomShape.line = undefined;
    generatedCustomShape.arrows = undefined;
    expect(generatedCustomShape.shadow).toMatchObject({
      kind: 'outer',
      rotateWithShape: true,
    });
    const reopened = await PptxDocument.open(await custom.document.write());
    const reopenedShape = reopened.slides[0]!.shapes[0] as ShapeModel;
    expect(reopenedShape.line).toBeUndefined();
    expect(reopenedShape.arrows).toBeUndefined();
    expect(reopenedShape.shadow).toMatchObject({ kind: 'outer', rotateWithShape: true });
  });

  it('compares shape line public output and strict native divergences', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const dashes = [
      'solid',
      'dash',
      'dashDot',
      'lgDash',
      'lgDashDot',
      'lgDashDotDot',
      'sysDash',
      'sysDot',
    ] as const;
    const generatedCases: readonly {
      readonly name: string;
      readonly line?: Record<string, unknown>;
    }[] = [
      { name: 'Line omitted' },
      { name: 'Line none', line: { type: 'none' } },
      { name: 'Line empty', line: {} },
      { name: 'Line missing color', line: { type: 'solid' } },
      { name: 'Line sRGB', line: { color: 'FF0000' } },
      { name: 'Line scheme', line: { color: generated.SchemeColor.accent2 } },
      { name: 'Line transparency', line: { color: '00FF00', transparency: 50 } },
      { name: 'Line zero transparency', line: { color: '0000FF', transparency: 0 } },
      { name: 'Line zero width', line: { color: '112233', width: 0 } },
      { name: 'Line positive width', line: { color: '223344', width: 2.5 } },
      { name: 'Line deprecated alpha', line: { color: '334455', alpha: 40 } },
      { name: 'Line deprecated dash', line: { color: '445566', lineDash: 'dash' } },
      ...dashes.map((dash) => ({
        name: `Line dash ${dash}`,
        line: { color: '556677', dashType: dash },
      })),
      {
        name: 'Line arrows',
        line: {
          color: '667788',
          width: 3,
          dashType: 'dashDot',
          beginArrowType: 'triangle',
          endArrowType: 'arrow',
        },
      },
    ];
    for (const fixture of generatedCases) {
      const options: Record<string, unknown> = { objectName: fixture.name };
      if (fixture.line !== undefined) options.line = fixture.line;
      generatedSlide.addShape(generated.ShapeType.rect!, options);
    }

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    expect(importedShapes.size).toBe(generatedCases.length);

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const nativeShapes = new Map<string, ShapeModel>([
      ['Line omitted', nativeSlide.addShape('rect', { name: 'Line omitted' })],
      ['Line none', nativeSlide.addShape('rect', {
        name: 'Line none',
        line: { kind: 'none' },
      })],
      ['Line sRGB', nativeSlide.addShape('rect', {
        name: 'Line sRGB',
        line: { kind: 'line', color: { kind: 'srgb', value: 'FF0000' } },
      })],
      ['Line scheme', nativeSlide.addShape('rect', {
        name: 'Line scheme',
        line: { kind: 'line', color: { kind: 'scheme', value: 'accent2' } },
      })],
      ['Line transparency', nativeSlide.addShape('rect', {
        name: 'Line transparency',
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '00FF00' },
          transparency: 50,
        },
      })],
      ['Line zero transparency', nativeSlide.addShape('rect', {
        name: 'Line zero transparency',
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '0000FF' },
          transparency: 0,
        },
      })],
      ['Line zero width', nativeSlide.addShape('rect', {
        name: 'Line zero width',
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '112233' },
          width: 0,
        },
      })],
      ['Line positive width', nativeSlide.addShape('rect', {
        name: 'Line positive width',
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '223344' },
          width: 2.5,
        },
      })],
      ...dashes.map((dash) => [
        `Line dash ${dash}`,
        nativeSlide.addShape('rect', {
          name: `Line dash ${dash}`,
          line: {
            kind: 'line',
            color: { kind: 'srgb', value: '556677' },
            dash,
          },
        }),
      ] as const),
      ['Line arrows', nativeSlide.addShape('rect', {
        name: 'Line arrows',
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '667788' },
          width: 3,
          dash: 'dashDot',
        },
      })],
    ] as const);

    for (const name of [
      'Line omitted',
      'Line sRGB',
      'Line scheme',
      'Line transparency',
      'Line positive width',
      ...dashes.map((dash) => `Line dash ${dash}`),
      'Line arrows',
    ]) {
      const importedShape = importedShapes.get(name)!;
      const nativeShape = nativeShapes.get(name)!;
      expect(importedShape, name).toBeInstanceOf(ShapeModel);
      expect(importedShape.name, name).toBe(nativeShape.name);
      expect(importedShape.presetType, name).toBe(nativeShape.presetType);
      expect(importedShape.transform, name).toEqual(nativeShape.transform);
      expect(importedShape.line, name).toEqual(nativeShape.line);
    }

    const generatedNone = importedShapes.get('Line none')!;
    const nativeNone = nativeShapes.get('Line none')!;
    expect(generatedNone.line).toBeUndefined();
    expect(nativeNone.line).toEqual({ kind: 'none' });
    expect(shapeXml(imported, 0, generatedNone.id)).toMatch(/<a:ln><\/a:ln>|<a:ln\/>/);
    expect(shapeXml(native, 0, nativeNone.id)).toContain('<a:ln><a:noFill/></a:ln>');

    for (const name of ['Line empty', 'Line missing color']) {
      expect(importedShapes.get(name)!.line, name).toEqual({
        kind: 'line',
        color: { kind: 'srgb', value: '333333' },
        width: 1,
        dash: 'solid',
      });
    }

    const generatedZeroTransparency = importedShapes.get('Line zero transparency')!;
    const nativeZeroTransparency = nativeShapes.get('Line zero transparency')!;
    expect(generatedZeroTransparency.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '0000FF' },
      width: 1,
      dash: 'solid',
    });
    expect(nativeZeroTransparency.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '0000FF' },
      transparency: 0,
      width: 1,
      dash: 'solid',
    });
    expect(shapeXml(imported, 0, generatedZeroTransparency.id)).not.toContain('<a:alpha');
    expect(shapeXml(native, 0, nativeZeroTransparency.id))
      .toContain('<a:alpha val="100000"/>');

    const generatedZeroWidth = importedShapes.get('Line zero width')!;
    const nativeZeroWidth = nativeShapes.get('Line zero width')!;
    expect(generatedZeroWidth.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      width: 1,
      dash: 'solid',
    });
    expect(nativeZeroWidth.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      width: 0,
      dash: 'solid',
    });
    expect(shapeXml(imported, 0, generatedZeroWidth.id)).toContain('<a:ln w="12700">');
    expect(shapeXml(native, 0, nativeZeroWidth.id)).toContain('<a:ln w="0">');

    expect(importedShapes.get('Line deprecated alpha')!.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '334455' },
      width: 1,
      dash: 'solid',
    });
    expect(importedShapes.get('Line deprecated dash')!.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '445566' },
      width: 1,
      dash: 'solid',
    });

    const arrows = importedShapes.get('Line arrows')!;
    expect(shapeXml(imported, 0, arrows.id)).toContain('<a:headEnd type="triangle"');
    expect(shapeXml(imported, 0, arrows.id)).toContain('<a:tailEnd type="arrow"');
    arrows.line = {
      kind: 'line',
      color: { kind: 'scheme', value: 'accent4' },
      transparency: 25,
      width: 2,
      dash: 'sysDash',
    };
    const editedArrowXml = shapeXml(imported, 0, arrows.id);
    expect(editedArrowXml).toContain('<a:headEnd type="triangle"');
    expect(editedArrowXml).toContain('<a:tailEnd type="arrow"');

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    for (const line of [
      {},
      { kind: 'line' },
      { type: 'none' },
      {
        kind: 'line',
        color: { kind: 'srgb', value: '334455' },
        alpha: 40,
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: '445566' },
        dashType: 'dash',
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: '445566' },
        lineDash: 'dash',
      },
    ]) {
      expect(() => nativeSlide.addShape('rect', { line } as never)).toThrow();
    }
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);

    const reopened = await PptxDocument.open(await imported.write());
    const reopenedArrows = reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'Line arrows',
    ) as ShapeModel;
    expect(reopenedArrows.line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent4' },
      transparency: 25,
      width: 2,
      dash: 'sysDash',
    });
    expect(shapeXml(reopened, 0, reopenedArrows.id)).toContain('<a:headEnd type="triangle"');
    expect(shapeXml(reopened, 0, reopenedArrows.id)).toContain('<a:tailEnd type="arrow"');
  });

  it('compares shape arrow public output and strict native divergences', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const arrowTypes = [
      'none',
      'arrow',
      'diamond',
      'oval',
      'stealth',
      'triangle',
    ] as const;
    const generatedCases: Record<string, Record<string, unknown>> = {};
    for (const type of arrowTypes) {
      generatedCases[`Arrow begin ${type}`] = {
        objectName: `Arrow begin ${type}`,
        line: { color: '112233', beginArrowType: type },
      };
      generatedCases[`Arrow end ${type}`] = {
        objectName: `Arrow end ${type}`,
        line: { color: '112233', endArrowType: type },
      };
    }
    Object.assign(generatedCases, {
      'Arrow both': {
        objectName: 'Arrow both',
        line: {
          color: '112233',
          beginArrowType: 'triangle',
          endArrowType: 'arrow',
        },
      },
      'Arrow only defaults': {
        objectName: 'Arrow only defaults',
        line: { beginArrowType: 'diamond' },
      },
      'Arrow none line': {
        objectName: 'Arrow none line',
        line: { type: 'none', beginArrowType: 'triangle' },
      },
      'Arrow empty ignored': {
        objectName: 'Arrow empty ignored',
        line: { color: '112233', beginArrowType: '', endArrowType: '' },
      },
      'Arrow nested aliases ignored': {
        objectName: 'Arrow nested aliases ignored',
        line: { color: '112233', lineHead: 'triangle', lineTail: 'arrow' },
      },
      'Arrow invalid passthrough': {
        objectName: 'Arrow invalid passthrough',
        line: { color: '112233', beginArrowType: 'bogus' },
      },
      'Arrow top aliases mapped': {
        objectName: 'Arrow top aliases mapped',
        line: { color: '112233' },
        lineHead: 'stealth',
        lineTail: 'oval',
      },
    });
    for (const options of Object.values(generatedCases)) {
      generatedSlide.addShape(generated.ShapeType.line!, options);
    }

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    expect(importedShapes.size).toBe(Object.keys(generatedCases).length);

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const nativeShapes = new Map<string, ShapeModel>();
    for (const type of arrowTypes) {
      nativeShapes.set(`Arrow begin ${type}`, nativeSlide.addShape('line', {
        name: `Arrow begin ${type}`,
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '112233' },
        },
        arrows: { begin: type },
      }));
      nativeShapes.set(`Arrow end ${type}`, nativeSlide.addShape('line', {
        name: `Arrow end ${type}`,
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '112233' },
        },
        arrows: { end: type },
      }));
    }
    nativeShapes.set('Arrow both', nativeSlide.addShape('line', {
      name: 'Arrow both',
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
      },
      arrows: { begin: 'triangle', end: 'arrow' },
    }));
    nativeShapes.set('Arrow only defaults', nativeSlide.addShape('line', {
      name: 'Arrow only defaults',
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: '333333' },
      },
      arrows: { begin: 'diamond' },
    }));
    nativeShapes.set('Arrow none line', nativeSlide.addShape('line', {
      name: 'Arrow none line',
      arrows: { begin: 'triangle' },
    }));
    nativeShapes.set('Arrow top aliases mapped', nativeSlide.addShape('line', {
      name: 'Arrow top aliases mapped',
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
      },
      arrows: { begin: 'stealth', end: 'oval' },
    }));

    for (const [name, nativeShape] of nativeShapes) {
      const importedShape = importedShapes.get(name)!;
      expect(importedShape, name).toBeInstanceOf(ShapeModel);
      expect(importedShape.name, name).toBe(nativeShape.name);
      expect(importedShape.presetType, name).toBe(nativeShape.presetType);
      expect(importedShape.transform, name).toEqual(nativeShape.transform);
      expect(importedShape.line, name).toEqual(nativeShape.line);
      expect(importedShape.arrows, name).toEqual(nativeShape.arrows);
      const xml = shapeXml(imported, 0, importedShape.id);
      if (importedShape.arrows?.begin !== undefined) {
        expect(xml, name).toContain(
          `<a:headEnd type="${importedShape.arrows.begin}"/>`,
        );
      }
      if (importedShape.arrows?.end !== undefined) {
        expect(xml, name).toContain(
          `<a:tailEnd type="${importedShape.arrows.end}"/>`,
        );
      }
      if (
        importedShape.arrows?.begin !== undefined
        && importedShape.arrows.end !== undefined
      ) {
        expect(xml.indexOf('<a:headEnd'), name).toBeLessThan(xml.indexOf('<a:tailEnd'));
      }
    }

    const generatedArrowOnly = importedShapes.get('Arrow only defaults')!;
    const nativeArrowOnly = nativeShapes.get('Arrow only defaults')!;
    expect(generatedArrowOnly.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '333333' },
      width: 1,
      dash: 'solid',
    });
    expect(nativeArrowOnly.line).toEqual(generatedArrowOnly.line);
    expect(shapeXml(imported, 0, generatedArrowOnly.id)).toContain('<a:ln w="12700">');

    const generatedNoneLine = importedShapes.get('Arrow none line')!;
    const nativeNoneLine = nativeShapes.get('Arrow none line')!;
    expect(generatedNoneLine.line).toBeUndefined();
    expect(nativeNoneLine.line).toBeUndefined();
    expect(generatedNoneLine.arrows).toEqual({ begin: 'triangle' });
    expect(nativeNoneLine.arrows).toEqual({ begin: 'triangle' });
    expect(shapeXml(imported, 0, generatedNoneLine.id))
      .toContain('<a:ln><a:headEnd type="triangle"/></a:ln>');
    expect(shapeXml(native, 0, nativeNoneLine.id))
      .toContain('<a:ln><a:headEnd type="triangle"/></a:ln>');
    const nativeExplicitNone = nativeSlide.addShape('line', {
      name: 'Arrow native explicit none',
      line: { kind: 'none' },
      arrows: { begin: 'triangle' },
    });
    expect(nativeExplicitNone.line).toEqual({ kind: 'none' });
    expect(nativeExplicitNone.arrows).toEqual({ begin: 'triangle' });
    expect(shapeXml(native, 0, nativeExplicitNone.id))
      .toContain('<a:ln><a:noFill/><a:headEnd type="triangle"/></a:ln>');

    for (const name of ['Arrow empty ignored', 'Arrow nested aliases ignored']) {
      expect(importedShapes.get(name)!.arrows, name).toBeUndefined();
      expect(shapeXml(imported, 0, importedShapes.get(name)!.id), name)
        .not.toMatch(/<a:(?:headEnd|tailEnd)\b/);
    }
    expect(importedShapes.get('Arrow invalid passthrough')!.arrows).toBeUndefined();
    expect(shapeXml(imported, 0, importedShapes.get('Arrow invalid passthrough')!.id))
      .toContain('<a:headEnd type="bogus"/>');
    expect(importedShapes.get('Arrow top aliases mapped')!.arrows)
      .toEqual({ begin: 'stealth', end: 'oval' });

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    for (const arrows of [
      { begin: '' },
      { begin: 'bogus' },
      { beginArrowType: 'arrow' },
      { endArrowType: 'arrow' },
      { lineHead: 'arrow' },
      { lineTail: 'arrow' },
    ]) {
      expect(() => nativeSlide.addShape('line', { arrows } as never)).toThrow(TypeError);
    }
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);

    const editable = importedShapes.get('Arrow both')!;
    const importedLine = editable.line;
    editable.arrows = { begin: 'diamond' };
    expect(editable.arrows).toEqual({ begin: 'diamond' });
    expect(editable.line).toEqual(importedLine);
    editable.arrows = undefined;
    expect(editable.arrows).toBeUndefined();
    expect(editable.line).toEqual(importedLine);
    editable.arrows = { end: 'stealth' };

    const reopened = await PptxDocument.open(await imported.write());
    const reopenedEditable = reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'Arrow both',
    ) as ShapeModel;
    expect(reopenedEditable.arrows).toEqual({ end: 'stealth' });
    expect(reopenedEditable.line).toEqual(importedLine);
    const reopenedXml = shapeXml(reopened, 0, reopenedEditable.id);
    expect(reopenedXml).not.toContain('<a:headEnd');
    expect(reopenedXml).toContain('<a:tailEnd type="stealth"/>');
  });

  it('locks every deprecated line alias against its owner-specific runtime behavior', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const dashes = [
      'solid',
      'dash',
      'dashDot',
      'lgDash',
      'lgDashDot',
      'lgDashDotDot',
      'sysDash',
      'sysDot',
    ] as const;
    const arrowTypes = [
      'none',
      'arrow',
      'diamond',
      'oval',
      'stealth',
      'triangle',
    ] as const;
    const baseLine = {
      kind: 'line' as const,
      color: { kind: 'srgb' as const, value: '112233' },
      width: 1,
      dash: 'solid' as const,
    };
    const expected = new Map<string, {
      readonly line: ShapeLine;
      readonly arrows?: ShapeArrows;
    }>();
    const addShapeCase = (
      name: string,
      options: Record<string, unknown>,
      line: ShapeLine,
      arrows?: ShapeArrows,
    ) => {
      generatedSlide.addShape(generated.ShapeType.line!, {
        objectName: name,
        ...options,
      });
      expected.set(name, arrows === undefined ? { line } : { line, arrows });
    };
    const addTextCase = (
      name: string,
      options: Record<string, unknown>,
      line: ShapeLine,
      arrows?: ShapeArrows,
    ) => {
      generatedSlide.addText(name, { objectName: name, ...options });
      expected.set(name, arrows === undefined ? { line } : { line, arrows });
    };

    for (const dash of dashes) {
      addShapeCase(`Nested shape lineDash ${dash}`, {
        line: { color: '112233', lineDash: dash },
      }, baseLine);
      addTextCase(`Nested text lineDash ${dash}`, {
        line: { color: '112233', lineDash: dash },
      }, baseLine);
      addShapeCase(`Top shape lineDash ${dash}`, {
        line: { color: '112233', dashType: 'solid' },
        lineDash: dash,
      }, { ...baseLine, dash });
      addTextCase(`Top text lineDash ignored ${dash}`, {
        line: { color: '112233', dashType: 'solid' },
        lineDash: dash,
      }, baseLine);
      addTextCase(`Top text-shape lineDash ${dash}`, {
        shape: generated.ShapeType.line,
        line: { color: '112233', dashType: 'solid' },
        lineDash: dash,
      }, { ...baseLine, dash });
    }

    for (const type of arrowTypes) {
      addShapeCase(`Nested shape lineHead lineTail ${type}`, {
        line: { color: '112233', lineHead: type, lineTail: type },
      }, baseLine);
      addTextCase(`Nested text lineHead lineTail ${type}`, {
        line: { color: '112233', lineHead: type, lineTail: type },
      }, baseLine);
      addShapeCase(`Top shape lineHead lineTail ${type}`, {
        line: { color: '112233' },
        lineHead: type,
        lineTail: type,
      }, baseLine, { begin: type, end: type });
      addTextCase(`Top text lineHead lineTail ignored ${type}`, {
        line: { color: '112233' },
        lineHead: type,
        lineTail: type,
      }, baseLine);
      addTextCase(`Top text-shape lineHead lineTail ${type}`, {
        shape: generated.ShapeType.line,
        line: { color: '112233' },
        lineHead: type,
        lineTail: type,
      }, baseLine, { begin: type, end: type });
    }

    for (const alias of ['pt', 'size'] as const) {
      addShapeCase(`Nested shape ${alias}`, {
        line: { color: '112233', [alias]: 2.5 },
      }, baseLine);
      addTextCase(`Nested text ${alias}`, {
        line: { color: '112233', [alias]: 2.5 },
      }, baseLine);
    }
    addShapeCase('Top shape lineSize', {
      line: { color: '112233', width: 1 },
      lineSize: 2.5,
    }, { ...baseLine, width: 2.5 });
    addTextCase('Top text lineSize ignored', {
      line: { color: '112233', width: 1 },
      lineSize: 2.5,
    }, baseLine);
    addTextCase('Top text-shape lineSize', {
      shape: generated.ShapeType.line,
      line: { color: '112233', width: 1 },
      lineSize: 2.5,
    }, { ...baseLine, width: 2.5 });
    addShapeCase('Nested shape alpha', {
      line: { color: '112233', alpha: 40 },
    }, baseLine);
    addTextCase('Nested text alpha', {
      line: { color: '112233', alpha: 40 },
    }, { ...baseLine, transparency: 40 });
    addTextCase('Nested text-shape alpha', {
      shape: generated.ShapeType.line,
      line: { color: '112233', alpha: 40 },
    }, baseLine);

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    expect(importedShapes.size).toBe(expected.size);
    for (const [name, state] of expected) {
      const shape = importedShapes.get(name);
      expect(shape, name).toBeInstanceOf(ShapeModel);
      expect(shape!.line, name).toEqual(state.line);
      expect(shape!.arrows, name).toEqual(state.arrows);
    }
  });

  it('compares text shape hyperlink public output and strict native divergences', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    generatedSlide.addText('Plain URL', {
      objectName: 'Plain URL',
      x: 1,
      y: 1,
      w: 2,
      h: 1,
      hyperlink: { url: 'https://example.com/path?a=1&b=2' },
    });
    generatedSlide.addText('First\nSecond', {
      objectName: 'Multiline URL',
      x: 1,
      y: 2,
      w: 2,
      h: 1,
      color: '123ABC',
      hyperlink: { url: 'mailto:test@example.com', tooltip: 'Mail & help' },
    });
    generatedSlide.addText('Internal slide', {
      objectName: 'Internal slide',
      x: 1,
      y: 3,
      w: 2,
      h: 1,
      hyperlink: { slide: 2 },
    });
    generatedSlide.addText([
      {
        text: 'Run one',
        options: { color: 'FF0000', hyperlink: { url: 'https://one.example/path' } },
      },
      {
        text: 'Run two',
        options: {
          color: '00AA00',
          underline: false,
          hyperlink: { slide: 2, tooltip: 'Run & two' },
        },
      },
    ], {
      objectName: 'Rich per-run',
      x: 1,
      y: 4,
      w: 3,
      h: 1,
    });
    generatedSlide.addText([
      { text: 'Same one', options: { hyperlink: { url: 'https://same.example/path' } } },
      { text: ' Same two', options: { hyperlink: { url: 'https://same.example/path' } } },
      { text: ' Plain' },
    ], {
      objectName: 'Rich identical targets',
      x: 4,
      y: 1,
      w: 3,
      h: 1,
    });
    generatedSlide.addText([
      { text: '', options: { hyperlink: { url: 'https://empty-run.example' } } },
      { text: 'Keep unlinked' },
    ], {
      objectName: 'Rich empty linked run',
      x: 4,
      y: 2,
      w: 3,
      h: 1,
    });
    generatedSlide.addText([
      { text: 'Inherited outer' },
      { text: ' Local valid', options: { hyperlink: { url: 'https://local-valid.example' } } },
    ], {
      objectName: 'Rich outer and local',
      x: 4,
      y: 3,
      w: 3,
      h: 1,
      hyperlink: { url: 'https://broken-outer.example' },
    });
    generated.addSlide();

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedSlide = imported.slides[0]!;
    const importedShapes = new Map(importedSlide.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    const clickAttributes = (xml: string) => [...xml.matchAll(
      /<a:hlinkClick\b([^>]*)>/g,
    )].map((match) => match[1]!);
    const clickIds = (xml: string) => clickAttributes(xml).map((attributes) =>
      attributes.match(/\br:id="([^"]+)"/)?.[1]);
    const nonVisualXml = (xml: string) => xml.slice(0, xml.indexOf('</p:nvSpPr>'));
    const relationship = (id: string | undefined) =>
      importedSlide.relationships.find((candidate) => candidate.id === id);

    const plain = importedShapes.get('Plain URL')!;
    const plainXml = shapeXml(imported, 0, plain.id);
    const plainClicks = clickAttributes(plainXml);
    const plainIds = clickIds(plainXml);
    expect(plain.hyperlink).toEqual({
      url: 'https://example.com/path?a=1&b=2',
      tooltip: '',
    });
    expect(plainClicks).toHaveLength(2);
    expect(new Set(plainIds).size).toBe(1);
    expect(clickAttributes(nonVisualXml(plainXml))).toHaveLength(1);
    expect(plainClicks.every((attributes) => attributes.includes('tooltip=""'))).toBe(true);
    expect(plainXml).toMatch(/<a:rPr\b[^>]*\bu="sng"[^>]*>[\s\S]*?<a:hlinkClick/);
    expect(relationship(plainIds[0])).toMatchObject({
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
      target: 'https://example.com/path?a=1&b=2',
      targetMode: 'External',
    });

    const multiline = importedShapes.get('Multiline URL')!;
    const multilineXml = shapeXml(imported, 0, multiline.id);
    const multilineClicks = clickAttributes(multilineXml);
    const multilineIds = clickIds(multilineXml);
    expect(multilineClicks).toHaveLength(3);
    expect(new Set(multilineIds).size).toBe(1);
    expect(clickAttributes(nonVisualXml(multilineXml))).toHaveLength(1);
    expect(multilineClicks.every(
      (attributes) => attributes.includes('tooltip="Mail &amp; help"'),
    )).toBe(true);
    expect(relationship(multilineIds[0])).toMatchObject({
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
      target: 'mailto:test@example.com',
      targetMode: 'External',
    });

    const internal = importedShapes.get('Internal slide')!;
    const internalXml = shapeXml(imported, 0, internal.id);
    const internalClicks = clickAttributes(internalXml);
    const internalIds = clickIds(internalXml);
    expect(internal.hyperlink).toEqual({ slide: 2, tooltip: '' });
    expect(internalClicks).toHaveLength(2);
    expect(new Set(internalIds).size).toBe(1);
    expect(internalClicks.every(
      (attributes) => attributes.includes('action="ppaction://hlinksldjump"'),
    )).toBe(true);
    expect(relationship(internalIds[0])).toMatchObject({
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
      target: 'slide2.xml',
      targetMode: 'Internal',
    });

    const richPerRun = importedShapes.get('Rich per-run')!;
    const richPerRunXml = shapeXml(imported, 0, richPerRun.id);
    const richPerRunClicks = clickAttributes(richPerRunXml);
    const richPerRunIds = clickIds(richPerRunXml);
    expect(richPerRun.hyperlink).toBeUndefined();
    expect(clickAttributes(nonVisualXml(richPerRunXml))).toEqual([]);
    expect(richPerRunClicks).toHaveLength(2);
    expect(new Set(richPerRunIds).size).toBe(2);
    expect(relationship(richPerRunIds[0])).toMatchObject({
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
      target: 'https://one.example/path',
      targetMode: 'External',
    });
    expect(relationship(richPerRunIds[1])).toMatchObject({
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
      target: 'slide2.xml',
      targetMode: 'Internal',
    });
    expect(richPerRunClicks[1]).toContain('tooltip="Run &amp; two"');
    expect(richPerRunClicks[1]).toContain('action="ppaction://hlinksldjump"');
    expect(richPerRunClicks[0]).toContain('action=""');
    expect(richPerRun.richText[0]!.runs.map((run) => ({
      hyperlink: run.style?.hyperlink,
      underline: run.style?.underline,
    }))).toEqual([
      {
        hyperlink: { url: 'https://one.example/path', tooltip: '' },
        underline: { style: 'sng' },
      },
      {
        hyperlink: { slide: 2, tooltip: 'Run & two' },
        underline: { style: 'sng' },
      },
    ]);
    expect(richPerRunXml).toContain('uri="{A12FA001-AC4F-418D-AE19-62706E023703}"');

    const identical = importedShapes.get('Rich identical targets')!;
    const identicalXml = shapeXml(imported, 0, identical.id);
    const identicalIds = clickIds(identicalXml);
    expect(identical.richText[0]!.runs.map((run) => run.style?.hyperlink)).toEqual([
      { url: 'https://same.example/path', tooltip: '' },
      { url: 'https://same.example/path', tooltip: '' },
      undefined,
    ]);
    expect(identicalIds).toHaveLength(2);
    expect(new Set(identicalIds).size).toBe(2);
    expect(identicalIds.map((id) => relationship(id)?.target)).toEqual([
      'https://same.example/path',
      'https://same.example/path',
    ]);

    const emptyLinked = importedShapes.get('Rich empty linked run')!;
    const emptyLinkedXml = shapeXml(imported, 0, emptyLinked.id);
    expect(emptyLinked.richText[0]!.runs.map((run) => ({
      text: run.text,
      hyperlink: run.style?.hyperlink,
    }))).toEqual([{ text: 'Keep unlinked', hyperlink: undefined }]);
    expect(clickIds(emptyLinkedXml)).toEqual([]);
    expect(importedSlide.relationships.find(
      ({ target }) => target === 'https://empty-run.example',
    )).toBeDefined();

    const outerAndLocal = importedShapes.get('Rich outer and local')!;
    const outerAndLocalXml = shapeXml(imported, 0, outerAndLocal.id);
    expect(outerAndLocal.hyperlink).toBeUndefined();
    expect(outerAndLocal.richText[0]!.runs.map((run) => run.style?.hyperlink)).toEqual([
      undefined,
      { url: 'https://local-valid.example', tooltip: '' },
    ]);
    expect(clickIds(outerAndLocalXml)).toContain('rIdundefined');
    expect(importedSlide.relationships.filter(
      ({ target }) => target === 'https://local-valid.example',
    )).toHaveLength(1);
    expect(importedSlide.relationships.some(
      ({ target }) => target === 'https://broken-outer.example',
    )).toBe(false);

    const brokenGenerated = new PptxGenJS();
    const brokenSlide = brokenGenerated.addSlide();
    brokenSlide.addText([
      { text: 'Rich outer one', options: { bold: true } },
      { text: ' two', options: { color: '00AA00', underline: false } },
    ], {
      objectName: 'Rich outer',
      x: 1,
      y: 1,
      w: 3,
      h: 1,
      hyperlink: { url: 'https://outer.example/path' },
    });
    const broken = await openPptxGenJSPublicOutput(brokenGenerated);
    const brokenShapeXml = [...slideXml(broken, 0).matchAll(
      /<p:sp(?:\s[^>]*)?>[\s\S]*?<\/p:sp>/g,
    )].map((match) => match[0]).find((xml) => xml.includes('name="Rich outer"'))!;
    const brokenClicks = clickAttributes(brokenShapeXml);
    expect(brokenClicks).toHaveLength(3);
    expect(brokenClicks.every(
      (attributes) => attributes.includes('r:id="rIdundefined"'),
    )).toBe(true);
    expect(broken.slides[0]!.relationships.filter(
      ({ type }) => type.endsWith('/hyperlink') || type.endsWith('/slide'),
    )).toEqual([]);

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    native.addSlide();
    const nativePlain = nativeSlide.addText('Plain URL', {
      name: 'Plain URL',
      x: inches(1),
      y: inches(1),
      width: inches(2),
      height: inches(1),
      hyperlink: { url: 'https://example.com/path?a=1&b=2' },
    });
    const nativeMultiline = nativeSlide.addText('First\nSecond', {
      name: 'Multiline URL',
      x: inches(1),
      y: inches(2),
      width: inches(2),
      height: inches(1),
      hyperlink: { url: 'mailto:test@example.com', tooltip: 'Mail & help' },
    });
    const nativeInternal = nativeSlide.addText('Internal slide', {
      name: 'Internal slide',
      x: inches(1),
      y: inches(3),
      width: inches(2),
      height: inches(1),
      hyperlink: { slide: 2 },
    });
    const nativeRichOuter = nativeSlide.addRichText([{
      runs: [
        { text: 'Rich outer one', style: { bold: true } },
        {
          text: ' two',
          style: {
            color: { kind: 'srgb', value: '00AA00' },
            underline: false,
          },
        },
      ],
    }], {
      name: 'Rich outer',
      x: inches(1),
      y: inches(4),
      width: inches(3),
      height: inches(1),
      hyperlink: { url: 'https://outer.example/path' },
    });
    const nativeRichPerRun = nativeSlide.addRichText([{
      runs: [
        {
          text: 'Run one',
          style: {
            color: { kind: 'srgb', value: 'FF0000' },
            hyperlink: { url: 'https://one.example/path' },
          },
        },
        {
          text: 'Run two',
          style: {
            color: { kind: 'srgb', value: '00AA00' },
            underline: false,
            hyperlink: { slide: 2, tooltip: 'Run & two' },
          },
        },
      ],
    }], { name: 'Native rich per-run' });
    const nativeIdentical = nativeSlide.addRichText([{
      runs: [
        { text: 'Same one', style: { hyperlink: { url: 'https://same.example/path' } } },
        { text: ' Same two', style: { hyperlink: { url: 'https://same.example/path' } } },
        { text: ' Plain' },
      ],
    }], { name: 'Native identical targets' });
    const nativeOuterAndLocal = nativeSlide.addRichText([{
      runs: [
        { text: 'Inherited outer' },
        {
          text: ' Local valid',
          style: { hyperlink: { url: 'https://local-valid.example' } },
        },
        { text: ' Suppressed', style: { hyperlink: false } },
      ],
    }], {
      name: 'Native outer and local',
      hyperlink: { url: 'https://valid-outer.example' },
    });

    for (const [generatedShape, nativeShape] of [
      [plain, nativePlain],
      [multiline, nativeMultiline],
      [internal, nativeInternal],
    ] as const) {
      expect(nativeShape.name).toBe(generatedShape.name);
      expect(nativeShape.text).toBe(generatedShape.text);
      expect(nativeShape.transform).toEqual(generatedShape.transform);
      const nativeXml = shapeXml(native, 0, nativeShape.id);
      expect(clickAttributes(nativeXml)).toHaveLength(clickAttributes(
        shapeXml(imported, 0, generatedShape.id),
      ).length);
      expect(new Set(clickIds(nativeXml)).size).toBe(1);
    }
    expect(nativePlain.hyperlink).toEqual({ url: 'https://example.com/path?a=1&b=2' });
    expect(nativeInternal.hyperlink).toEqual({ slide: 2 });
    expect(nativeRichOuter.hyperlink).toEqual({ url: 'https://outer.example/path' });
    const nativeRichOuterXml = shapeXml(native, 0, nativeRichOuter.id);
    const nativeRichOuterIds = clickIds(nativeRichOuterXml);
    expect(nativeRichOuterIds).toHaveLength(3);
    expect(new Set(nativeRichOuterIds).size).toBe(1);
    expect(nativeRichOuterXml).not.toContain('rIdundefined');
    expect(nativeRichOuterXml).toMatch(/<a:rPr\b[^>]*\bu="sng"[^>]*>/);
    expect(nativeRichOuterXml).toMatch(/<a:rPr\b[^>]*\bu="none"[^>]*>/);
    expect(nativeSlide.relationships.find(
      ({ id }) => id === nativeRichOuterIds[0],
    )).toMatchObject({
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
      target: 'https://outer.example/path',
      targetMode: 'External',
    });
    expect(nativeRichPerRun.richText[0]!.runs.map((run) => ({
      hyperlink: run.style?.hyperlink,
      underline: run.style?.underline,
    }))).toEqual([
      { hyperlink: { url: 'https://one.example/path' }, underline: { style: 'sng' } },
      { hyperlink: { slide: 2, tooltip: 'Run & two' }, underline: false },
    ]);
    const nativeIdenticalIds = clickIds(shapeXml(native, 0, nativeIdentical.id));
    expect(nativeIdenticalIds).toHaveLength(2);
    expect(new Set(nativeIdenticalIds).size).toBe(2);
    expect(nativeOuterAndLocal.hyperlink).toEqual({ url: 'https://valid-outer.example' });
    expect(nativeOuterAndLocal.richText[0]!.runs.map((run) => run.style?.hyperlink)).toEqual([
      { url: 'https://valid-outer.example' },
      { url: 'https://local-valid.example' },
      undefined,
    ]);
    expect(shapeXml(native, 0, nativeOuterAndLocal.id)).not.toContain('rIdundefined');

    const beforeInvalid = packageState(native);
    const shapeCount = nativeSlide.shapes.length;
    for (const hyperlink of [
      {},
      { url: '' },
      { url: 'https://example.com', slide: 2 },
      { slide: 99 },
      { url: 42 },
      { slide: '2' },
    ]) {
      expect(() => nativeSlide.addText('Invalid text hyperlink', {
        hyperlink,
      } as never)).toThrow();
      expect(() => nativeSlide.addRichText([{ runs: [{ text: 'Invalid rich link' }] }], {
        hyperlink,
      } as never)).toThrow();
      expect(packageState(native)).toEqual(beforeInvalid);
      expect(nativeSlide.shapes).toHaveLength(shapeCount);
    }
    expect(() => nativeSlide.addRichText([{
      runs: [{ text: '', style: { hyperlink: { url: 'https://empty-run.example' } } }],
    }])).toThrow(/non-empty text/i);
    expect(packageState(native)).toEqual(beforeInvalid);
    expect(nativeSlide.shapes).toHaveLength(shapeCount);
  });

  it('compares shape hyperlink public output and strict native divergences', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const cases: readonly {
      readonly name: string;
      readonly hyperlink: unknown;
    }[] = [
      { name: 'URL', hyperlink: { url: 'https://example.com?a=1&b=2' } },
      {
        name: 'URL tooltip',
        hyperlink: { url: 'mailto:test@example.com', tooltip: 'Mail & help' },
      },
      { name: 'Slide', hyperlink: { slide: 2 } },
      { name: 'Slide tooltip', hyperlink: { slide: 3, tooltip: '' } },
      { name: 'Self', hyperlink: { slide: 1 } },
      { name: 'Empty', hyperlink: {} },
      { name: 'Both', hyperlink: { url: 'https://example.com', slide: 2 } },
      { name: 'Zero', hyperlink: { slide: 0 } },
      { name: 'Negative', hyperlink: { slide: -1 } },
      { name: 'Fraction', hyperlink: { slide: 1.5 } },
      { name: 'Out of range', hyperlink: { slide: 99 } },
      { name: 'Numeric URL', hyperlink: { url: 42 } },
      { name: 'String value', hyperlink: 'https://example.com' },
    ];
    const consoleOutput = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    let generatedBytes: Uint8Array;
    try {
      for (const { name, hyperlink } of cases) {
        generatedSlide.addShape(generated.ShapeType.rect!, {
          objectName: name,
          x: 1,
          y: 1,
          w: 1,
          h: 1,
          hyperlink,
        });
      }
      generated.addSlide();
      generated.addSlide();
      generatedBytes = await generated.write({
        outputType: 'nodebuffer',
        compression: true,
      });
      expect(consoleOutput).toHaveBeenCalledTimes(3);
      expect(consoleOutput.mock.calls.map(([message]) => String(message)).join('\n'))
        .toContain('hyperlink requires either');
      expect(consoleOutput.mock.calls.map(([message]) => String(message)).join('\n'))
        .toContain('should be an object');
    } finally {
      consoleOutput.mockRestore();
    }

    const imported = await PptxDocument.open(generatedBytes!);
    const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    expect(importedShapes.get('URL')!.hyperlink).toEqual({
      url: 'https://example.com?a=1&b=2',
      tooltip: '',
    });
    expect(importedShapes.get('URL tooltip')!.hyperlink).toEqual({
      url: 'mailto:test@example.com',
      tooltip: 'Mail & help',
    });
    expect(importedShapes.get('Slide')!.hyperlink).toEqual({ slide: 2, tooltip: '' });
    expect(importedShapes.get('Slide tooltip')!.hyperlink)
      .toEqual({ slide: 3, tooltip: '' });
    expect(importedShapes.get('Self')!.hyperlink).toEqual({ slide: 1, tooltip: '' });
    expect(importedShapes.get('Numeric URL')!.hyperlink)
      .toEqual({ url: '42', tooltip: '' });
    for (const name of [
      'Empty',
      'Both',
      'Zero',
      'Negative',
      'Fraction',
      'Out of range',
      'String value',
    ]) {
      expect(importedShapes.get(name)!.hyperlink, name).toBeUndefined();
    }

    const firstSlideXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    expect(shapeXml(imported, 0, importedShapes.get('URL')!.id))
      .toContain('tooltip=""');
    expect(shapeXml(imported, 0, importedShapes.get('URL')!.id))
      .not.toContain('ppaction://hlinksldjump');
    expect(shapeXml(imported, 0, importedShapes.get('Slide')!.id))
      .toContain('action="ppaction://hlinksldjump"');
    expect(shapeXml(imported, 0, importedShapes.get('Both')!.id).match(/<a:hlinkClick/g))
      .toHaveLength(2);
    for (const name of ['Empty', 'Zero', 'String value']) {
      expect(shapeXml(imported, 0, importedShapes.get(name)!.id), name)
        .not.toContain('<a:hlinkClick');
    }
    expect(firstSlideXml).toContain('tooltip="Mail &amp; help"');

    const generatedRelationships = imported.slides[0]!.relationships.filter(
      ({ type }) => type.endsWith('/hyperlink') || type.endsWith('/slide'),
    );
    expect(generatedRelationships.map(({ type, target, targetMode }) => ({
      type: type.slice(type.lastIndexOf('/') + 1),
      target,
      targetMode,
    }))).toEqual([
      { type: 'hyperlink', target: 'https://example.com?a=1&b=2', targetMode: 'External' },
      { type: 'hyperlink', target: 'mailto:test@example.com', targetMode: 'External' },
      { type: 'slide', target: 'slide2.xml', targetMode: 'Internal' },
      { type: 'slide', target: 'slide3.xml', targetMode: 'Internal' },
      { type: 'slide', target: 'slide1.xml', targetMode: 'Internal' },
      { type: 'slide', target: 'slidehttps://example.com.xml', targetMode: 'Internal' },
      { type: 'slide', target: 'slide-1.xml', targetMode: 'Internal' },
      { type: 'slide', target: 'slide1.5.xml', targetMode: 'Internal' },
      { type: 'slide', target: 'slide99.xml', targetMode: 'Internal' },
      { type: 'hyperlink', target: '42', targetMode: 'External' },
    ]);
    for (const target of [
      '/ppt/slides/slidehttps:/example.com.xml',
      '/ppt/slides/slide-1.xml',
      '/ppt/slides/slide1.5.xml',
      '/ppt/slides/slide99.xml',
    ]) {
      expect(imported.opcPackage.hasPart(target), target).toBe(false);
    }

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    native.addSlide();
    native.addSlide();
    const nativeShapes = [
      nativeSlide.addShape('rect', {
        name: 'URL',
        hyperlink: { url: 'https://example.com?a=1&b=2' },
      }),
      nativeSlide.addShape('rect', {
        name: 'URL tooltip',
        hyperlink: { url: 'mailto:test@example.com', tooltip: 'Mail & help' },
      }),
      nativeSlide.addShape('rect', { name: 'Slide', hyperlink: { slide: 2 } }),
      nativeSlide.addShape('rect', {
        name: 'Slide tooltip',
        hyperlink: { slide: 3, tooltip: '' },
      }),
      nativeSlide.addShape('rect', { name: 'Self', hyperlink: { slide: 1 } }),
    ];
    expect(nativeShapes.map(({ hyperlink }) => hyperlink)).toEqual([
      { url: 'https://example.com?a=1&b=2' },
      { url: 'mailto:test@example.com', tooltip: 'Mail & help' },
      { slide: 2 },
      { slide: 3, tooltip: '' },
      { slide: 1 },
    ]);
    expect(nativeSlide.relationships.filter(
      ({ type }) => type.endsWith('/hyperlink') || type.endsWith('/slide'),
    ).map(({ type, target, targetMode }) => ({
      type: type.slice(type.lastIndexOf('/') + 1),
      target,
      targetMode,
    }))).toEqual(generatedRelationships.slice(0, 5).map(({ type, target, targetMode }) => ({
      type: type.slice(type.lastIndexOf('/') + 1),
      target,
      targetMode,
    })));

    const beforeInvalid = {
      parts: native.opcPackage.parts.map(({ uri, contentType, bytes, relationships }) => ({
        uri,
        contentType,
        bytes: bytes.slice(),
        relationships,
      })),
      mutations: [...native.opcPackage.mutations],
      shapes: [...nativeSlide.shapes],
    };
    let accessorCalls = 0;
    const accessors = (['url', 'slide', 'tooltip'] as const).map((key) =>
      Object.defineProperty({}, key, {
        enumerable: true,
        get() {
          accessorCalls += 1;
          return key === 'slide' ? 1 : 'https://example.com';
        },
      }));
    for (const hyperlink of [
      {},
      { url: '' },
      { url: 42 },
      { url: 'https://example.com', slide: 2 },
      { slide: 0 },
      { slide: -1 },
      { slide: 1.5 },
      { slide: 99 },
      'https://example.com',
      { url: 'https://example.com', _rId: 'rId9' },
      { target: 'https://example.com' },
      { kind: 'url', url: 'https://example.com' },
      { url: 'https://example.com', [Symbol('unsafe')]: true },
      ...accessors,
    ]) {
      expect(() => nativeSlide.addShape('rect', { hyperlink } as never)).toThrow();
    }
    expect(accessorCalls).toBe(0);
    expect({
      parts: native.opcPackage.parts.map(({ uri, contentType, bytes, relationships }) => ({
        uri,
        contentType,
        bytes,
        relationships,
      })),
      mutations: native.opcPackage.mutations,
      shapes: nativeSlide.shapes,
    }).toEqual(beforeInvalid);

    const validGenerated = new PptxGenJS();
    const validSlide = validGenerated.addSlide();
    for (const { name, hyperlink } of cases.slice(0, 5)) {
      validSlide.addShape(validGenerated.ShapeType.rect!, {
        objectName: name,
        x: 1,
        y: 1,
        w: 1,
        h: 1,
        hyperlink,
      });
    }
    validGenerated.addSlide();
    validGenerated.addSlide();
    const validImported = await openPptxGenJSPublicOutput(validGenerated);
    const duplicate = validImported.duplicateSlide(0);
    expect((duplicate.shapes[4] as ShapeModel).hyperlink).toEqual({
      slide: validImported.slides.indexOf(duplicate) + 1,
      tooltip: '',
    });
    validImported.moveSlide(1, 0);
    expect((validImported.slides[1]!.shapes[2] as ShapeModel).hyperlink)
      .toEqual({ slide: 1, tooltip: '' });
    expect((validImported.slides[1]!.shapes[4] as ShapeModel).hyperlink)
      .toEqual({ slide: 2, tooltip: '' });
    const reopened = await PptxDocument.open(await validImported.write());
    expect((reopened.slides[1]!.shapes[0] as ShapeModel).hyperlink)
      .toEqual({ url: 'https://example.com?a=1&b=2', tooltip: '' });
    expect((reopened.slides[3]!.shapes[4] as ShapeModel).hyperlink)
      .toEqual({ slide: 4, tooltip: '' });
  });

  it('compares shape shadow public output and strict native divergences', async () => {
    expect(new PptxGenJS().version).toBe('4.0.1');
    const generate = async (name: string, shadow: unknown) => {
      const generated = new PptxGenJS();
      const slide = generated.addSlide();
      const options: Record<string, unknown> = {
        objectName: name,
        x: 1,
        y: 1,
        w: 2,
        h: 1,
      };
      if (shadow !== undefined) options.shadow = shadow;
      slide.addShape(generated.ShapeType.roundRect!, options);
      const bytes = await generated.write({
        outputType: 'nodebuffer',
        compression: true,
      });
      const document = await PptxDocument.open(bytes);
      const partUri = document.slides[0]!.partUri;
      const xml = new TextDecoder().decode(document.opcPackage.requirePart(partUri).bytes);
      return { document, xml };
    };

    const omitted = await generate('Omitted', undefined);
    const none = await generate('None', { type: 'none' });
    for (const fixture of [omitted, none]) {
      expect(fixture.xml).not.toContain('<a:effectLst>');
      expect(fixture.xml).not.toMatch(/<(?:a:)?(?:inner|outer)Shdw/);
      expect((fixture.document.slides[0]!.shapes[0] as ShapeModel).shadow).toBeUndefined();
    }

    const defaults = await generate('Outer defaults', { type: 'outer' });
    expect(defaults.xml).toContain(
      '<a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="bl" ' +
      'rotWithShape="0" blurRad="101600" dist="50800" dir="16200000"> ' +
      '<a:srgbClr val="000000"> <a:alpha val="75000"/></a:srgbClr> ' +
      '</a:outerShdw>',
    );
    expect((defaults.document.slides[0]!.shapes[0] as ShapeModel).shadow).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '000000' },
      opacity: 0.75,
      blur: 8,
      angle: 270,
      distance: 4,
      rotateWithShape: false,
    });

    const custom = await generate('Outer custom', {
      type: 'outer',
      color: '123ABC',
      opacity: 0.42,
      blur: 7.25,
      angle: 123.4,
      offset: 5.5,
      rotateWithShape: true,
    });
    expect(custom.xml).toContain(
      'rotWithShape="0" blurRad="92075" dist="69850" dir="7404000"> ' +
      '<a:srgbClr val="123ABC"> <a:alpha val="42000"/>',
    );
    const generatedCustom = (custom.document.slides[0]!.shapes[0] as ShapeModel).shadow;
    expect(generatedCustom).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '123ABC' },
      opacity: 0.42,
      blur: 7.25,
      angle: 123.4,
      distance: 5.5,
      rotateWithShape: false,
    });

    const zero = await generate('Outer zero', {
      type: 'outer',
      color: '000000',
      opacity: 0,
      blur: 0,
      angle: 0,
      offset: 0,
      rotateWithShape: false,
    });
    expect((zero.document.slides[0]!.shapes[0] as ShapeModel).shadow).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '000000' },
      opacity: 0.75,
      blur: 8,
      angle: 270,
      distance: 4,
      rotateWithShape: false,
    });

    const malformedInner = await generate('Inner', { type: 'inner' });
    expect(malformedInner.xml).toContain('<a:innerShdw  blurRad="101600"');
    expect(malformedInner.xml).toContain('</a:outerShdw></a:effectLst>');
    expect(() => malformedInner.document.slides[0]!.shapes).toThrow();
    await expect(malformedInner.document.write()).rejects.toThrow();

    const hashColor = await generate('Hash color', {
      type: 'outer',
      color: '#ABCDEF',
    });
    expect(hashColor.xml).toContain('<a:srgbClr val="#ABCDEF">');
    expect((hashColor.document.slides[0]!.shapes[0] as ShapeModel).shadow).toBeUndefined();

    const unknownType = await generate('Unknown type', {
      type: 'bogus',
      color: 'FF0000',
    });
    expect(unknownType.xml).toContain('<a:bogusShdw  blurRad="101600"');
    expect(unknownType.xml).toContain('</a:outerShdw></a:effectLst>');
    expect(() => unknownType.document.slides[0]!.shapes).toThrow();

    const invalidRanges = await generate('Invalid ranges', {
      type: 'outer',
      color: '00FF00',
      opacity: 2,
      blur: -1,
      angle: 400,
      offset: 201,
    });
    expect(invalidRanges.xml).toContain(
      'blurRad="-12700" dist="2552700" dir="24000000"> ' +
      '<a:srgbClr val="00FF00"> <a:alpha val="200000"/>',
    );
    expect((invalidRanges.document.slides[0]!.shapes[0] as ShapeModel).shadow)
      .toBeUndefined();

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const nativeDefaults = nativeSlide.addShape('roundRect', {
      shadow: { kind: 'outer' },
    });
    const nativeCustom = nativeSlide.addShape('roundRect', {
      shadow: {
        kind: 'outer',
        color: { kind: 'srgb', value: '123ABC' },
        opacity: 0.42,
        blur: 7.25,
        angle: 123.4,
        distance: 5.5,
        rotateWithShape: true,
      },
    });
    const nativeZero = nativeSlide.addShape('roundRect', {
      shadow: {
        kind: 'outer',
        color: { kind: 'scheme', value: 'accent2' },
        opacity: 0,
        blur: 0,
        angle: 0,
        distance: 0,
        rotateWithShape: false,
      },
    });
    const nativeInner = nativeSlide.addShape('roundRect', {
      shadow: { kind: 'inner', color: { kind: 'scheme', value: 'accent3' } },
    });
    expect(nativeDefaults.shadow).toEqual(
      (defaults.document.slides[0]!.shapes[0] as ShapeModel).shadow,
    );
    expect(nativeCustom.shadow).toMatchObject({
      kind: generatedCustom?.kind,
      color: generatedCustom?.color,
      opacity: generatedCustom?.opacity,
      blur: generatedCustom?.blur,
      angle: generatedCustom?.angle,
      distance: generatedCustom?.distance,
      rotateWithShape: true,
    });
    expect(generatedCustom?.kind === 'outer' && generatedCustom.rotateWithShape).toBe(false);
    expect(nativeZero.shadow).toEqual({
      kind: 'outer',
      color: { kind: 'scheme', value: 'accent2' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
      rotateWithShape: false,
    });
    expect(nativeInner.shadow?.kind).toBe('inner');
    expect(shapeXml(native, 0, nativeInner.id)).toContain(
      '<a:innerShdw blurRad="101600" dist="50800" dir="16200000">',
    );
    expect(shapeXml(native, 0, nativeInner.id)).toContain('</a:innerShdw>');
    const reopenedNative = await PptxDocument.open(await native.write());
    expect((reopenedNative.slides[0]!.shapes[3] as ShapeModel).shadow)
      .toEqual(nativeInner.shadow);

    const beforeInvalid = {
      parts: native.opcPackage.parts.map(({ uri, contentType, bytes, relationships }) => ({
        uri,
        contentType,
        bytes: bytes.slice(),
        relationships,
      })),
      mutations: [...native.opcPackage.mutations],
      shapes: [...nativeSlide.shapes],
    };
    let accessorCalls = 0;
    const accessor = Object.defineProperty({}, 'kind', {
      enumerable: true,
      get() {
        accessorCalls += 1;
        return 'outer';
      },
    });
    for (const shadow of [
      {},
      { kind: 'none' },
      { type: 'outer' },
      { kind: 'outer', offset: 4 },
      { kind: 'inner', rotateWithShape: false },
      { kind: 'outer', opacity: 2 },
      { kind: 'outer', blur: -1 },
      { kind: 'outer', angle: 400 },
      { kind: 'outer', distance: 201 },
      { kind: 'outer', color: { kind: 'srgb', value: '#ABC' } },
      accessor,
    ]) {
      expect(() => nativeSlide.addShape('rect', { shadow } as never)).toThrow();
    }
    expect(accessorCalls).toBe(0);
    expect({
      parts: native.opcPackage.parts.map(({ uri, contentType, bytes, relationships }) => ({
        uri,
        contentType,
        bytes,
        relationships,
      })),
      mutations: native.opcPackage.mutations,
      shapes: nativeSlide.shapes,
    }).toEqual(beforeInvalid);
  });

  it('locks ShadowProps across shape, text, image, and chart owners against PptxGenJS 4.0.1', async () => {
    const atomIds = [
      ...[
        'angle',
        'blur',
        'color',
        'offset',
        'opacity',
        'rotateWithShape',
        'type',
      ].map((property) => `interface:ShadowProps@property:${property}`),
      ...['inner', 'none', 'outer']
        .map((value) => `union:interface:ShadowProps@property:type#${value}`),
      'interface:ShapeProps@property:shadow',
      'interface:TextPropsOptions@property:shadow',
    ];
    expect(new Set(atomIds).size).toBe(12);
    expect(atomIds).not.toContain('interface:ImageProps@property:shadow');

    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const slide = generated.addSlide();
    const outer = () => ({
      type: 'outer' as const,
      color: '123456',
      opacity: 0.6,
      blur: 4,
      angle: 135,
      offset: 5,
      rotateWithShape: true,
    });
    slide.addShape(generated.ShapeType.rect!, {
      objectName: 'Shape shared shadow',
      x: 0,
      y: 0,
      w: 1,
      h: 1,
      shadow: outer(),
    });
    slide.addText('Text shared shadow', {
      objectName: 'Text shared shadow',
      x: 1,
      y: 0,
      w: 1,
      h: 1,
      shadow: outer(),
    });
    slide.addShape(generated.ShapeType.rect!, {
      objectName: 'Shape shared shadow none',
      x: 2,
      y: 0,
      w: 1,
      h: 1,
      shadow: { type: 'none' },
    });
    slide.addText('Text shared shadow none', {
      objectName: 'Text shared shadow none',
      x: 3,
      y: 0,
      w: 1,
      h: 1,
      shadow: { type: 'none' },
    });
    slide.addImage({
      data: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
      x: 0,
      y: 1,
      w: 1,
      h: 1,
      shadow: {
        type: 'inner',
        color: '112233',
        opacity: 0.4,
        blur: 2,
        angle: 45,
        offset: 3,
        rotateWithShape: true,
      },
    });
    slide.addChart(generated.ChartType.bar!, [{
      name: 'Series',
      labels: ['A'],
      values: [1],
    }], {
      x: 1,
      y: 1,
      w: 2,
      h: 2,
      shadow: {
        type: 'inner',
        color: '223344',
        opacity: 0.5,
        blur: 3,
        angle: 90,
        offset: 4,
        rotateWithShape: true,
      },
    });

    const document = await PptxDocument.open(await generated.write({
      outputType: 'nodebuffer',
      compression: true,
    }));
    const slideXml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    const chartXml = document.opcPackage.parts
      .filter(({ uri }) => uri.startsWith('/ppt/charts/chart') && uri.endsWith('.xml'))
      .map(({ bytes }) => new TextDecoder().decode(bytes))
      .join('');

    const slideOuterTags = slideXml.match(/<a:outerShdw\b[^>]*>/gu) ?? [];
    expect(slideOuterTags).toHaveLength(2);
    for (const tag of slideOuterTags) {
      expect(tag).toContain('rotWithShape="0"');
      expect(tag).toContain('blurRad="50800"');
      expect(tag).toContain('dist="63500"');
      expect(tag).toContain('dir="8100000"');
    }
    expect(slideXml.match(/<\/a:outerShdw>/gu) ?? []).toHaveLength(2);
    expect(slideXml.match(/<a:srgbClr val="123456">/gu) ?? []).toHaveLength(2);
    expect(slideXml.match(/<a:alpha val="60000"\/>/gu) ?? []).toHaveLength(2);
    const noneOwners = document.slides[0]!.shapes.filter(({ name }) =>
      name === 'Shape shared shadow none' || name === 'Text shared shadow none');
    expect(noneOwners).toHaveLength(2);
    for (const owner of noneOwners) {
      expect((owner as ShapeModel).shadow).toBeUndefined();
      expect(shapeXml(document, 0, owner.id)).not.toContain('<a:effectLst>');
    }
    const imageInnerTags = slideXml.match(/<a:innerShdw\b[^>]*>/gu) ?? [];
    expect(imageInnerTags).toHaveLength(1);
    expect(imageInnerTags[0]).toContain('blurRad="25400"');
    expect(imageInnerTags[0]).toContain('dist="38100"');
    expect(imageInnerTags[0]).toContain('dir="2700000"');
    expect(imageInnerTags[0]).not.toContain('rotWithShape');
    expect(slideXml.match(/<\/a:innerShdw>/gu) ?? []).toHaveLength(1);
    expect(slideXml).toContain('<a:srgbClr val="112233">');
    expect(slideXml).toContain('<a:alpha val="40000"/>');

    const chartInnerTags = chartXml.match(/<a:innerShdw\b[^>]*>/gu) ?? [];
    expect(chartInnerTags).toHaveLength(1);
    expect(chartInnerTags[0]).toContain('blurRad="38100"');
    expect(chartInnerTags[0]).toContain('rotWithShape="1"');
    expect(chartInnerTags[0]).toContain('dist="50800"');
    expect(chartInnerTags[0]).toContain('dir="5400000"');
    expect(chartXml.match(/<\/a:innerShdw>/gu) ?? []).toHaveLength(1);
    expect(chartXml).toContain('<a:srgbClr val="223344">');
    expect(chartXml).toContain('<a:alpha val="50000"/>');
    expect(document.diagnostics).toEqual([]);
  });

  it('reads and round-trips PptxGenJS preset shape adjustment output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const fixtures = [
      {
        name: 'Round radius',
        type: 'roundRect',
        options: { rectRadius: 0.5, w: 4, h: 2 },
        expected: [{ name: 'adj', value: 25_000 }],
      },
      {
        name: 'Pie angles',
        type: 'pie',
        options: { angleRange: [270, 0] },
        expected: [
          { name: 'adj1', value: 16_200_000 },
          { name: 'adj2', value: 0 },
        ],
      },
      {
        name: 'Fractional arc angles',
        type: 'arc',
        options: { angleRange: [12.34567, 89.99999] },
        expected: [
          { name: 'adj1', value: 740_740 },
          { name: 'adj2', value: 5_399_999 },
        ],
      },
      {
        name: 'Block arc angles and thickness',
        type: 'blockArc',
        options: { angleRange: [270, 0], arcThicknessRatio: 0.5 },
        expected: [
          { name: 'adj1', value: 16_200_000 },
          { name: 'adj2', value: 0 },
          { name: 'adj3', value: 25_000 },
        ],
      },
      {
        name: 'Omitted shortcuts',
        type: 'roundRect',
        options: {},
        expected: [],
      },
      {
        name: 'Zero radius shortcut',
        type: 'roundRect',
        options: { rectRadius: 0, w: 4, h: 2 },
        expected: [],
      },
      {
        name: 'Zero thickness shortcut',
        type: 'blockArc',
        options: { angleRange: [270, 0], arcThicknessRatio: 0 },
        expected: [
          { name: 'adj1', value: 16_200_000 },
          { name: 'adj2', value: 0 },
        ],
      },
    ] as const;

    for (const fixture of fixtures) {
      generatedSlide.addShape(generated.ShapeType[fixture.type]!, {
        objectName: fixture.name,
        ...fixture.options,
      });
    }

    const imported = await importPptxGenJS(generated);
    const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    for (const fixture of fixtures) {
      const importedShape = importedShapes.get(fixture.name);
      expect(importedShape, fixture.name).toBeInstanceOf(ShapeModel);
      expect(importedShape?.adjustments, fixture.name).toEqual(fixture.expected);
      const nativeShape = nativeSlide.addShape(fixture.type, {
        name: fixture.name,
        adjustments: fixture.expected,
      });
      expect(nativeShape.adjustments, fixture.name).toEqual(importedShape?.adjustments);
    }

    const reopenedImported = await PptxDocument.open(await imported.write());
    const reopenedNative = await PptxDocument.open(await native.write());
    const expected = fixtures.map((fixture) => fixture.expected);
    expect(reopenedImported.slides[0]!.shapes.map((shape) =>
      (shape as ShapeModel).adjustments)).toEqual(expected);
    expect(reopenedNative.slides[0]!.shapes.map((shape) =>
      (shape as ShapeModel).adjustments)).toEqual(expected);
  });

  it('records PptxGenJS adjustment shortcut divergences from native lists', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const generatedCases = [
      {
        name: 'Dropped zero radius',
        type: 'roundRect',
        options: { rectRadius: 0, w: 4, h: 2 },
      },
      {
        name: 'Dropped zero thickness',
        type: 'blockArc',
        options: { angleRange: [270, 0], arcThicknessRatio: 0 },
      },
      {
        name: 'String coercion',
        type: 'blockArc',
        options: { angleRange: ['12.5', '90'], arcThicknessRatio: '0.25' },
      },
      {
        name: 'Radius precedence',
        type: 'blockArc',
        options: {
          rectRadius: 0.5,
          angleRange: [270, 0],
          arcThicknessRatio: 0.5,
          w: 4,
          h: 2,
        },
      },
      {
        name: 'Thickness without angles',
        type: 'blockArc',
        options: { arcThicknessRatio: 0.5 },
      },
      {
        name: 'Out of range angles',
        type: 'arc',
        options: { angleRange: [-1, 361] },
      },
      {
        name: 'Malformed coerced angles',
        type: 'arc',
        options: { angleRange: ['not-an-angle', 'also-not-an-angle'] },
      },
      {
        name: 'Unsafe coerced angles',
        type: 'arc',
        options: { angleRange: [Number.MAX_SAFE_INTEGER, Number.MAX_SAFE_INTEGER] },
      },
    ] as const;
    for (const fixture of generatedCases) {
      generatedSlide.addShape(generated.ShapeType[fixture.type]!, {
        objectName: fixture.name,
        ...fixture.options,
      });
    }

    const imported = await importPptxGenJS(generated);
    const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    expect(importedShapes.get('Dropped zero radius')?.adjustments).toEqual([]);
    expect(importedShapes.get('Dropped zero thickness')?.adjustments).toEqual([
      { name: 'adj1', value: 16_200_000 },
      { name: 'adj2', value: 0 },
    ]);
    expect(importedShapes.get('String coercion')?.adjustments).toEqual([
      { name: 'adj1', value: 750_000 },
      { name: 'adj2', value: 5_400_000 },
      { name: 'adj3', value: 12_500 },
    ]);
    expect(importedShapes.get('Radius precedence')?.adjustments).toEqual([
      { name: 'adj', value: 25_000 },
    ]);
    expect(importedShapes.get('Thickness without angles')?.adjustments).toEqual([]);
    expect(importedShapes.get('Out of range angles')?.adjustments).toEqual([
      { name: 'adj1', value: -60_000 },
      { name: 'adj2', value: 60_000 },
    ]);
    const malformed = importedShapes.get('Malformed coerced angles')!;
    expect(malformed.adjustments).toBeUndefined();
    const malformedXml = shapeXml(imported, 0, malformed.id);
    expect(malformedXml).toContain('fmla="val NaN"');
    const unsafe = importedShapes.get('Unsafe coerced angles')!;
    expect(unsafe.adjustments).toBeUndefined();
    const unsafeXml = shapeXml(imported, 0, unsafe.id);
    expect(unsafeXml).toContain('fmla="val 540431955284437800000"');
    const reopened = await PptxDocument.open(await imported.write());
    const reopenedMalformed = reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'Malformed coerced angles',
    ) as ShapeModel;
    const reopenedUnsafe = reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'Unsafe coerced angles',
    ) as ShapeModel;
    expect(reopenedMalformed.adjustments).toBeUndefined();
    expect(reopenedUnsafe.adjustments).toBeUndefined();
    expect(shapeXml(reopened, 0, reopenedMalformed.id)).toBe(malformedXml);
    expect(shapeXml(reopened, 0, reopenedUnsafe.id)).toBe(unsafeXml);

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    expect(nativeSlide.addShape('roundRect', {
      adjustments: [{ name: 'adj', value: 0 }],
    }).adjustments).toEqual([{ name: 'adj', value: 0 }]);
    expect(() => nativeSlide.addShape('roundRect', {
      adjustments: [{ name: 'adj', value: '0' }],
    } as never)).toThrow(TypeError);
    const deliberateFinalList = [
      { name: 'adj', value: 25_000 },
      { name: 'adj1', value: 16_200_000 },
      { name: 'adj2', value: 0 },
      { name: 'adj3', value: 25_000 },
    ];
    expect(nativeSlide.addShape('blockArc', {
      adjustments: deliberateFinalList,
    }).adjustments).toEqual(deliberateFinalList);
  });

  it('keeps custom geometry connection sites at the PptxGenJS public boundary', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.addSlide().addShape(generated.ShapeType.custGeom!, {
      objectName: 'Public custom geometry',
      x: 1,
      y: 1,
      w: 4,
      h: 3,
      points: [{ x: 0, y: 0 }, { x: 4, y: 3 }, { close: true }],
    });

    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShape = imported.slides[0]!.shapes[0] as ShapeModel;
    expect(importedShape.name).toBe('Public custom geometry');
    expect(importedShape.customGeometry).toBeDefined();
    expect(Object.hasOwn(importedShape.customGeometry!, 'connectionSites')).toBe(false);
    expect(Object.hasOwn(importedShape.customGeometry!, 'textRectangle')).toBe(false);
    expect(shapeXml(imported, 0, importedShape.id)).toMatch(
      /<a:cxnLst(?:\s*\/>|\s*>\s*<\/a:cxnLst>)/,
    );
    expect(shapeXml(imported, 0, importedShape.id)).toMatch(
      /<a:rect\s+l="l"\s+t="t"\s+r="r"\s+b="b"\s*\/>/,
    );

    const adapterSource = await readFile(new URL('./index.ts', import.meta.url), 'utf8');
    expect(adapterSource).toContain('presentation.write({');
    expect(adapterSource).not.toMatch(/\.\s*_[A-Za-z]/);
    expect(adapterSource).not.toMatch(/\[\s*['"]_[^'"]*['"]\s*\]/);
  });

  it('imports every legal PptxGenJS custom path command as native geometry', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    expect(generated.ShapeType.custGeom).toBe('custGeom');
    const points: readonly PptxGenJSCustomPoint[] = [
      { x: 0, y: 0 },
      { x: 1, y: 0 },
      { x: 2, y: 1, curve: { type: 'quadratic', x1: 1.5, y1: 0 } },
      {
        x: 3,
        y: 2,
        curve: { type: 'cubic', x1: 2.25, y1: 1, x2: 2.75, y2: 2 },
      },
      {
        x: 999,
        y: 999,
        curve: { type: 'arc', wR: 1, hR: 0.5, stAng: 30, swAng: 120 },
      },
      { x: 0.5, y: 0.5, moveTo: true },
      { x: 1.25, y: 1.25 },
      { close: true },
    ];
    generated.addSlide().addShape(generated.ShapeType.custGeom!, {
      objectName: 'All custom commands',
      x: 1,
      y: 1,
      w: 4,
      h: 3,
      points,
    });

    const expected: CustomGeometry = {
      paths: [{
        width: inches(4),
        height: inches(3),
        commands: [
          { kind: 'moveTo', point: { x: 0, y: 0 } },
          { kind: 'lineTo', point: { x: inches(1), y: 0 } },
          {
            kind: 'quadraticBezierTo',
            control: { x: inches(1.5), y: 0 },
            end: { x: inches(2), y: inches(1) },
          },
          {
            kind: 'cubicBezierTo',
            control1: { x: inches(2.25), y: inches(1) },
            control2: { x: inches(2.75), y: inches(2) },
            end: { x: inches(3), y: inches(2) },
          },
          {
            kind: 'arcTo',
            widthRadius: inches(1),
            heightRadius: inches(0.5),
            startAngle: degrees(30),
            sweepAngle: degrees(120),
          },
          { kind: 'moveTo', point: { x: inches(0.5), y: inches(0.5) } },
          { kind: 'lineTo', point: { x: inches(1.25), y: inches(1.25) } },
          { kind: 'close' },
        ],
      }],
    };
    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShape = imported.slides[0]!.shapes[0] as ShapeModel;
    expect(importedShape).toBeInstanceOf(ShapeModel);
    expect(importedShape.name).toBe('All custom commands');
    expect(importedShape.presetType).toBeUndefined();
    expect(importedShape.transform).toEqual({
      x: inches(1),
      y: inches(1),
      width: inches(4),
      height: inches(3),
      rotation: 0,
      flipHorizontal: false,
      flipVertical: false,
    });
    expect(importedShape.customGeometry).toEqual(expected);
    expect(Object.hasOwn(importedShape.customGeometry!, 'adjustments')).toBe(false);
    expect(Object.hasOwn(importedShape.customGeometry!, 'guides')).toBe(false);
    expect(Object.hasOwn(importedShape.customGeometry!, 'handles')).toBe(false);
    expect(shapeXml(imported, 0, importedShape.id)).not.toContain('x="999"');
    expect(shapeXml(imported, 0, importedShape.id)).not.toContain('y="999"');
    const beforeEvaluation = imported.opcPackage.requirePart(
      imported.slides[0]!.partUri,
    ).bytes.slice();
    const evaluationJournal = [...imported.opcPackage.mutations];
    const evaluated = importedShape.evaluateCustomGeometry();
    expect(evaluated?.paths).toEqual(importedShape.customGeometry?.paths);
    expect(evaluated?.textRectangle).toEqual({
      left: 0,
      top: 0,
      right: inches(4),
      bottom: inches(3),
    });
    expect(Object.isFrozen(evaluated)).toBe(true);
    expect(Object.isFrozen(evaluated?.paths[0]?.commands[2])).toBe(true);
    expect(imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes)
      .toEqual(beforeEvaluation);
    expect(imported.opcPackage.mutations).toEqual(evaluationJournal);

    const native = PptxDocument.create();
    const nativeShape = native.addSlide().addCustomShape(expected, {
      name: 'All custom commands',
      x: inches(1),
      y: inches(1),
      width: inches(4),
      height: inches(3),
    });
    expect(nativeShape.customGeometry).toEqual(importedShape.customGeometry);
    const reopened = await PptxDocument.open(await imported.write());
    expect((reopened.slides[0]!.shapes[0] as ShapeModel).customGeometry).toEqual(expected);
  });

  it('classifies PptxGenJS custom path unit heuristics and malformed runtime output', async () => {
    const generated = new PptxGenJS();
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    const validCases: readonly {
      readonly name: string;
      readonly points: readonly PptxGenJSCustomPoint[];
      readonly expected: CustomGeometry;
    }[] = [
      {
        name: 'Custom empty',
        points: [],
        expected: { paths: [{ width: inches(4), height: inches(3), commands: [] }] },
      },
      {
        name: 'Custom direct numeric',
        points: [{ x: 100, y: 200 }, { x: 300, y: 400 }],
        expected: {
          paths: [{
            width: inches(4),
            height: inches(3),
            commands: [
              { kind: 'moveTo', point: { x: 100, y: 200 } },
              { kind: 'lineTo', point: { x: 300, y: 400 } },
            ],
          }],
        },
      },
      {
        name: 'Custom numeric strings',
        points: [{ x: '1', y: '2' }, { x: '100', y: '200' }],
        expected: {
          paths: [{
            width: inches(4),
            height: inches(3),
            commands: [
              { kind: 'moveTo', point: { x: inches(1), y: inches(2) } },
              { kind: 'lineTo', point: { x: 100, y: 200 } },
            ],
          }],
        },
      },
      {
        name: 'Custom percentages',
        points: [{ x: '10%', y: '20%' }, { x: '50%', y: '60%' }],
        expected: {
          paths: [{
            width: inches(4),
            height: inches(3),
            commands: [
              { kind: 'moveTo', point: { x: 1_219_200, y: 1_371_600 } },
              { kind: 'lineTo', point: { x: 6_096_000, y: 4_114_800 } },
            ],
          }],
        },
      },
      {
        name: 'Custom later move',
        points: [
          { x: 0, y: 0 },
          { x: 1, y: 1 },
          { x: 2, y: 2, moveTo: true },
          { x: 3, y: 3 },
        ],
        expected: {
          paths: [{
            width: inches(4),
            height: inches(3),
            commands: [
              { kind: 'moveTo', point: { x: 0, y: 0 } },
              { kind: 'lineTo', point: { x: inches(1), y: inches(1) } },
              { kind: 'moveTo', point: { x: inches(2), y: inches(2) } },
              { kind: 'lineTo', point: { x: inches(3), y: inches(3) } },
            ],
          }],
        },
      },
      {
        name: 'Custom invalid kind omitted',
        points: [
          { x: 0, y: 0 },
          { x: 1, y: 1, curve: { type: 'bogus', x1: 0, y1: 0 } } as never,
        ],
        expected: {
          paths: [{
            width: inches(4),
            height: inches(3),
            commands: [{ kind: 'moveTo', point: { x: 0, y: 0 } }],
          }],
        },
      },
      {
        name: 'Custom missing cubic defaults',
        points: [
          { x: 0, y: 0 },
          { x: 1, y: 1, curve: { type: 'cubic', x1: 0, y1: 0 } } as never,
        ],
        expected: {
          paths: [{
            width: inches(4),
            height: inches(3),
            commands: [
              { kind: 'moveTo', point: { x: 0, y: 0 } },
              {
                kind: 'cubicBezierTo',
                control1: { x: 0, y: 0 },
                control2: { x: 0, y: 0 },
                end: { x: inches(1), y: inches(1) },
              },
            ],
          }],
        },
      },
    ];
    for (const fixture of validCases) {
      slide.addShape(generated.ShapeType.custGeom!, {
        objectName: fixture.name,
        x: 1,
        y: 1,
        w: 4,
        h: 3,
        points: fixture.points,
      });
    }
    const imported = await openPptxGenJSPublicOutput(generated);
    const importedShapes = new Map(imported.slides[0]!.shapes.map((shape) => [
      shape.name,
      shape as ShapeModel,
    ]));
    for (const fixture of validCases) {
      const snapshot = importedShapes.get(fixture.name)?.customGeometry;
      expect(snapshot, fixture.name).toEqual(fixture.expected);
      expect(Object.hasOwn(snapshot!, 'adjustments'), fixture.name).toBe(false);
      expect(Object.hasOwn(snapshot!, 'guides'), fixture.name).toBe(false);
      expect(Object.hasOwn(snapshot!, 'handles'), fixture.name).toBe(false);
    }
    expect(imported.slideSize).toEqual({ width: 12_192_000, height: 6_858_000 });

    const unsupported = new PptxGenJS();
    unsupported.layout = 'LAYOUT_WIDE';
    const unsupportedSlide = unsupported.addSlide();
    const unsupportedCases: readonly {
      readonly name: string;
      readonly points: readonly PptxGenJSCustomPoint[];
      readonly malformedXml: string;
    }[] = [
      {
        name: 'Custom first arc',
        points: [{
          x: 0,
          y: 0,
          curve: { type: 'arc', wR: 1, hR: 1, stAng: 0, swAng: 90 },
        }],
        malformedXml: '<a:arcTo hR="914400" wR="914400"',
      },
      {
        name: 'Custom zero radius',
        points: [
          { x: 0, y: 0 },
          { x: 1, y: 1, curve: { type: 'arc', wR: 0, hR: 1, stAng: 0, swAng: 90 } },
        ],
        malformedXml: 'wR="0"',
      },
      {
        name: 'Custom negative radius',
        points: [
          { x: 0, y: 0 },
          { x: 1, y: 1, curve: { type: 'arc', wR: -1, hR: 1, stAng: 0, swAng: 90 } },
        ],
        malformedXml: 'wR="-914400"',
      },
      {
        name: 'Custom unsafe coordinate',
        points: [{ x: 0, y: 0 }, { x: Number.MAX_SAFE_INTEGER + 1, y: 1 }],
        malformedXml: 'x="9007199254740992"',
      },
    ];
    for (const fixture of unsupportedCases) {
      unsupportedSlide.addShape(unsupported.ShapeType.custGeom!, {
        objectName: fixture.name,
        x: 1,
        y: 1,
        w: 4,
        h: 3,
        points: fixture.points,
      });
    }
    const unsupportedImported = await openPptxGenJSPublicOutput(unsupported);
    const unsupportedShapes = unsupportedImported.slides[0]!.shapes as ShapeModel[];
    const unsupportedXml = unsupportedShapes.map((shape) =>
      shapeXml(unsupportedImported, 0, shape.id));
    for (const [index, fixture] of unsupportedCases.entries()) {
      expect(unsupportedShapes[index]?.name).toBe(fixture.name);
      expect(unsupportedShapes[index]?.customGeometry, fixture.name).toBeUndefined();
      expect(unsupportedXml[index], fixture.name).toContain(fixture.malformedXml);
    }
    const unsupportedReopened = await PptxDocument.open(await unsupportedImported.write());
    for (const [index, source] of unsupportedXml.entries()) {
      const reopenedShape = unsupportedReopened.slides[0]!.shapes[index] as ShapeModel;
      expect(reopenedShape.customGeometry).toBeUndefined();
      expect(shapeXml(unsupportedReopened, 0, reopenedShape.id)).toBe(source);
    }

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const directNative = nativeSlide.addCustomShape({
      paths: [{
        width: 1,
        height: 1,
        commands: [{ kind: 'moveTo', point: { x: 1, y: 2 } }],
      }],
    });
    expect(directNative.customGeometry?.paths[0]?.commands[0]).toEqual({
      kind: 'moveTo',
      point: { x: 1, y: 2 },
    });
    // PptxGenJS 4.0.1 has no public guide-formula input; this is a native extension.
    const formulaNative = nativeSlide.addCustomShape({
      adjustments: [{ name: 'adj1', formula: { operator: 'val', operands: [25_000] } }],
      guides: [{ name: 'x1', formula: { operator: '*/', operands: ['w', 'adj1', 100_000] } }],
      paths: [{
        width: 100_000,
        height: 100_000,
        commands: [{ kind: 'moveTo', point: { x: 'x1', y: 0 } }],
      }],
    });
    expect(formulaNative.customGeometry).toEqual({
      adjustments: [{ name: 'adj1', formula: { operator: 'val', operands: [25_000] } }],
      guides: [{ name: 'x1', formula: { operator: '*/', operands: ['w', 'adj1', 100_000] } }],
      paths: [{
        width: 100_000,
        height: 100_000,
        commands: [{ kind: 'moveTo', point: { x: 'x1', y: 0 } }],
      }],
    });
    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    for (const geometry of [
      {
        paths: [{
          width: 1,
          height: 1,
          commands: [{ kind: 'moveTo', point: { x: '10 %', y: 0 } }],
        }],
      },
      {
        paths: [{
          width: 1,
          height: 1,
          commands: [{ kind: 'arcTo', widthRadius: 1, heightRadius: 1, startAngle: 0, sweepAngle: 1 }],
        }],
      },
      {
        paths: [{
          width: 1,
          height: 1,
          commands: [
            { kind: 'moveTo', point: { x: 0, y: 0 } },
            { kind: 'arcTo', widthRadius: 0, heightRadius: 1, startAngle: 0, sweepAngle: 1 },
          ],
        }],
      },
      {
        paths: [{
          width: 1,
          height: 1,
          commands: [
            { kind: 'moveTo', point: { x: 0, y: 0 } },
            { kind: 'unknown' },
          ],
        }],
      },
      {
        paths: [{
          width: 1,
          height: 1,
          commands: [
            { kind: 'moveTo', point: { x: 0, y: 0 } },
            { kind: 'lineTo', point: { x: Number.MAX_SAFE_INTEGER + 1, y: 0 } },
          ],
        }],
      },
    ]) expect(() => nativeSlide.addCustomShape(geometry as never)).toThrow();
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
  });

  it('reads every legal PptxGenJS preset shape public output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const legalPublicTypes = PRESET_SHAPE_TYPES.filter((type) => type !== 'foldedCorner');
    const slide = generated.addSlide();
    for (const type of legalPublicTypes) {
      const publicType = generated.ShapeType[type];
      expect(publicType, type).toBe(type);
      slide.addShape(publicType!);
    }

    const imported = await openPptxGenJSPublicOutput(generated);
    expect(imported.slides[0]?.shapes).toHaveLength(177);
    expect(imported.slides[0]?.shapes.map((shape) => {
      expect(shape).toBeInstanceOf(ShapeModel);
      return (shape as ShapeModel).presetType;
    })).toEqual(legalPublicTypes);
  }, 30_000);

  it('isolates the folderCorner defect from valid preset shape public output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    expect(generated.ShapeType.folderCorner).toBe('folderCorner');
    expect(generated.ShapeType.foldedCorner).toBeUndefined();
    expect(generated.ShapeType.custGeom).toBe('custGeom');
    generated.addSlide().addShape(generated.ShapeType.folderCorner!);

    const imported = await openPptxGenJSPublicOutput(generated);
    const malformedShape = imported.slides[0]?.shapes[0];
    expect(malformedShape).toBeInstanceOf(ShapeModel);
    expect((malformedShape as ShapeModel).presetType).toBeUndefined();
    expect(shapeXml(imported, 0, malformedShape!.id)).toContain('prst="folderCorner"');

    expect(PRESET_SHAPE_TYPES).toContain('foldedCorner');
    expect(PRESET_SHAPE_TYPES).not.toContain('folderCorner');
    expect(PRESET_SHAPE_TYPES).not.toContain('custGeom');
    const native = PptxDocument.create();
    const foldedCorner = native.addSlide().addShape('foldedCorner');
    const reopened = await PptxDocument.open(await native.write());
    expect(foldedCorner.presetType).toBe('foldedCorner');
    expect((reopened.slides[0]?.shapes[0] as ShapeModel).presetType).toBe('foldedCorner');
    expect(() => native.slides[0]!.addShape('folderCorner' as never)).toThrow(TypeError);
  });

  it('matches native basic table creation to public PptxGenJS plain-table output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const generatedSlide = generated.addSlide();
    const rows = [
      ['A1', 'B1', 'C1'],
      ['A2', 'B2', 'C2'],
    ] as const;
    generatedSlide.addTable(
      rows.map((row) => row.map((text) => ({ text, options: {} }))),
      { x: 1, y: 1.5, w: 6, h: 2, colW: [1, 2, 3], rowH: [0.75, 1.25] },
    );

    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes[0] as TableModel;
    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeOptions = {
      x: inches(1),
      y: inches(1.5),
      width: inches(6),
      height: inches(2),
      columnWidths: [inches(1), inches(2), inches(3)],
      rowHeights: [inches(0.75), inches(1.25)],
    };
    const nativeTable = native.addSlide().addTable(
      rows.map((row) => row.map((text) => ({ text }))),
      nativeOptions,
    );
    const stringNative = PptxDocument.create({ slideSize: 'wide' });
    stringNative.addSlide().addTable(rows, nativeOptions);

    expect(importedTable).toBeInstanceOf(TableModel);
    expect(nativeTable.transform).toMatchObject(importedTable.transform);
    expect(nativeTable.rows.map(({ cells }) => cells.map(({ text }) => text))).toEqual(
      importedTable.rows.map(({ cells }) => cells.map(({ text }) => text)),
    );
    expect(nativeTable.rows.map(({ cells }) => cells.map(({ margins }) => margins))).toEqual(
      importedTable.rows.map(({ cells }) => cells.map(({ margins }) => margins)),
    );
    expect(nativeTable.rows.map(({ cells }) => cells.map(({ borders }) => borders))).toEqual(
      importedTable.rows.map(({ cells }) => cells.map(({ borders }) => borders)),
    );

    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(native.slides[0]!.partUri).bytes,
    );
    const stringNativeXml = new TextDecoder().decode(
      stringNative.opcPackage.requirePart(stringNative.slides[0]!.partUri).bytes,
    );
    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    expect(nativeXml).toBe(stringNativeXml);
    for (const xml of [nativeXml, importedXml]) {
      expect(xml).toContain(
        'uri="http://schemas.openxmlformats.org/drawingml/2006/table"',
      );
      const columnWidths = [...xml.matchAll(/<a:gridCol w="(\d+)"\/>/g)]
        .map((match) => Number(match[1]));
      const rowHeights = [...xml.matchAll(/<a:tr h="(\d+)">/g)]
        .map((match) => Number(match[1]));
      expect(columnWidths).toHaveLength(3);
      expect(columnWidths).toEqual([inches(1), inches(2), inches(3)]);
      expect(columnWidths.reduce((sum, width) => sum + width, 0)).toBe(5_486_400);
      expect(rowHeights).toEqual([inches(0.75), inches(1.25)]);
      expect(xml.match(/<a:tc>/g)).toHaveLength(6);
      expect(xml.match(/marL="91440" marR="91440" marT="45720" marB="45720"/g))
        .toHaveLength(6);
      const properties = xml.match(/<a:tcPr[^>]*>([\s\S]*?)<\/a:tcPr>/)?.[1];
      expect(properties).toBeDefined();
      const left = properties!.indexOf('<a:lnL ');
      const right = properties!.indexOf('<a:lnR ');
      const top = properties!.indexOf('<a:lnT ');
      const bottom = properties!.indexOf('<a:lnB ');
      expect(left).toBeGreaterThanOrEqual(0);
      expect(left).toBeLessThan(right);
      expect(right).toBeLessThan(top);
      expect(top).toBeLessThan(bottom);
    }
    expect(nativeXml).not.toContain('p14:modId');
    expect(nativeXml).toContain('<a:ext cx="5486400" cy="1828800"/>');

    expect(importedTable.columnWidths).toEqual([
      inches(1),
      inches(2),
      inches(3),
    ]);
    importedTable.setColumnWidths([
      inches(1.5),
      inches(1.5),
      inches(3),
    ]);
    expect(importedTable.columnWidths).toEqual([
      inches(1.5),
      inches(1.5),
      inches(3),
    ]);
    expect(importedTable.transform.width).toBe(inches(6));
    expect(importedTable.rowHeights).toEqual([
      inches(0.75),
      inches(1.25),
    ]);
    importedTable.setRowHeights([inches(1), inches(1.5)]);
    expect(importedTable.rowHeights).toEqual([
      inches(1),
      inches(1.5),
    ]);
    expect(importedTable.transform.height).toBe(inches(2.5));

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    const reopenedNativeTable = reopenedNative.slides[0]!.shapes[0] as TableModel;
    const reopenedImportedTable = reopenedImported.slides[0]!.shapes[0] as TableModel;
    expect(reopenedNativeTable.rows).toEqual(nativeTable.rows);
    expect(reopenedImportedTable.rows).toEqual(importedTable.rows);
    expect(reopenedNativeTable.transform).toEqual(nativeTable.transform);
    expect(reopenedImportedTable.transform).toEqual(importedTable.transform);
    expect(reopenedImportedTable.columnWidths).toEqual([
      inches(1.5),
      inches(1.5),
      inches(3),
    ]);
    expect(reopenedImportedTable.rowHeights).toEqual([
      inches(1),
      inches(1.5),
    ]);
  });

  it('imports PptxGenJS 4.0.1 auto-page output with repeated headers and editable tables', async () => {
    const noOverflow = new PptxGenJS();
    const noOverflowSource = noOverflow.addSlide();
    noOverflowSource.addTable([['Header'], ['Only body']], {
      x: 1,
      y: 1,
      w: 8,
      autoPage: true,
      autoPageRepeatHeader: true,
      autoPageHeaderRows: 1,
    });
    expect(noOverflow.version).toBe('4.0.1');
    expect(noOverflowSource.newAutoPagedSlides).toEqual([]);
    const noOverflowDocument = await importPptxGenJS(noOverflow);
    expect((noOverflowDocument.slides[0]!.shapes[0] as TableModel).rows.map(
      (row) => row.cells[0]!.text,
    )).toEqual(['Header', 'Only body']);

    const generated = new PptxGenJS();
    const source = generated.addSlide();
    const following = generated.addSlide();
    following.addText('FOLLOWING', {
      x: 1,
      y: 1,
      w: 3,
      h: 0.5,
      objectName: 'Following sentinel',
    });
    const generatedUrl: { url: string; tooltip: string; _rId?: number } = {
      url: 'https://auto-page.example?a=1&b=2',
      tooltip: 'Auto-page URL',
    };
    const generatedInternal: { slide: number; tooltip: string; _rId?: number } = {
      slide: 2,
      tooltip: 'Existing following slide',
    };
    const generatedRows: (string | PptxGenJSTableCell)[][] = [
      ['Header A', 'Header B'],
      [
        { text: 'URL', options: { hyperlink: generatedUrl } },
        { text: 'Internal', options: { hyperlink: generatedInternal } },
      ],
      [
        {
          text: [
            { text: 'Rich', options: { bold: true, color: 'FF0000' } },
            { text: 'body', options: { italic: true } },
          ],
        },
        'Rich value',
      ],
      ...Array.from({ length: 12 }, (_, index) => [
        `Body ${index}`,
        `Value ${index}`,
      ]),
    ];
    const generatedOptions: Record<string, unknown> = {
      x: 1,
      y: 4,
      w: 8,
      autoPage: true,
      autoPageRepeatHeader: true,
      autoPageHeaderRows: 1,
      autoPageSlideStartY: 1.25,
      slideMargin: [0.5, 0.5, 0.75, 0.5],
      fontSize: 18,
      margin: 0.05,
    };
    source.addTable(generatedRows, generatedOptions);

    expect(generated.version).toBe('4.0.1');
    expect(source.newAutoPagedSlides).toHaveLength(2);
    expect(source.newAutoPagedSlides[0]).toBe(following);
    expect(generatedOptions).toMatchObject({
      autoPage: false,
      y: inches(1.25),
      margin: [0.05, 0.05, 0.05, 0.05],
      autoPageLineWeight: 0,
    });
    expect(generatedOptions).toHaveProperty('_arrObjTabHeadRows');
    expect(generatedUrl._rId).toBeTypeOf('number');
    expect(generatedInternal._rId).toBeTypeOf('number');

    const document = await importPptxGenJS(generated);
    const tables = document.slides.map((slide) => slide.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    )!);
    expect(tables).toHaveLength(3);
    expect(tables.every((table) => table instanceof TableModel)).toBe(true);
    expect(tables.map(({ transform }) => transform.y))
      .toEqual([inches(4), inches(1.25), inches(1.25)]);
    expect(tables.map((table) => table.rows[0]!.cells.map(({ text }) => text)))
      .toEqual(Array(3).fill(['Header A', 'Header B']));
    expect(tables.flatMap((table) => table.rows.slice(1).map(
      (row) => row.cells[0]!.text,
    ))).toEqual([
      'URL',
      'Richbody',
      ...Array.from({ length: 12 }, (_, index) => `Body ${index}`),
    ]);
    const richCell = tables.flatMap((table) => table.rows.flatMap(({ cells }) => cells))
      .find(({ text }) => text === 'Richbody')!;
    expect(richCell.richText[0]!.runs.map(({ text, style }) => ({
      text,
      bold: style?.bold,
      italic: style?.italic,
      color: style?.color,
    }))).toEqual([
      {
        text: 'Rich',
        bold: true,
        italic: undefined,
        color: { kind: 'srgb', value: 'FF0000' },
      },
      {
        text: 'body',
        bold: undefined,
        italic: true,
        color: { kind: 'srgb', value: '000000' },
      },
    ]);
    expect(tables[0]!.rows[1]!.cells.map(({ hyperlink }) => hyperlink)).toEqual([
      { url: 'https://auto-page.example?a=1&b=2', tooltip: 'Auto-page URL' },
      { slide: 2, tooltip: 'Existing following slide' },
    ]);
    expect(document.slides[1]!.shapes.map(({ name }) => name))
      .toContain('Following sentinel');
    expect(document.slides[2]!.shapes.map(({ name }) => name))
      .not.toContain('Following sentinel');

    tables.forEach((table, index) => table.setCellRichText(0, 1, [{
      runs: [{ text: `Editable header ${index}` }],
    }]));
    const reopened = await PptxDocument.open(await document.write());
    const reopenedTables = reopened.slides.map((slide) => slide.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    )!);
    expect(reopenedTables.map((table) => table.rows[0]!.cells[1]!.text)).toEqual([
      'Editable header 0',
      'Editable header 1',
      'Editable header 2',
    ]);
    expect(reopenedTables[0]!.rows[1]!.cells.map(({ hyperlink }) => hyperlink)).toEqual([
      { url: 'https://auto-page.example?a=1&b=2', tooltip: 'Auto-page URL' },
      { slide: 2, tooltip: 'Existing following slide' },
    ]);

    const twoHeaders = new PptxGenJS();
    const twoHeaderSource = twoHeaders.addSlide();
    twoHeaderSource.addTable([
      ['Header 1'],
      ['Header 2'],
      ...Array.from({ length: 15 }, (_, index) => [`Body ${index}`]),
    ], {
      x: 1,
      y: 4,
      w: 8,
      autoPage: true,
      autoPageRepeatHeader: true,
      autoPageHeaderRows: 2,
      autoPageSlideStartY: 1,
      fontSize: 18,
      margin: 0.05,
    });
    expect(twoHeaderSource.newAutoPagedSlides).toHaveLength(2);
    const twoHeaderDocument = await importPptxGenJS(twoHeaders);
    const twoHeaderTables = twoHeaderDocument.slides.map((slide) => slide.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    )!);
    expect(twoHeaderTables.map((table) => table.rows.slice(0, 2).map(
      (row) => row.cells[0]!.text,
    ))).toEqual(Array(3).fill(['Header 1', 'Header 2']));
    expect(twoHeaderTables.flatMap((table) => table.rows.slice(2).map(
      (row) => row.cells[0]!.text,
    ))).toEqual(Array.from({ length: 15 }, (_, index) => `Body ${index}`));
  }, 60_000);

  it('keeps strict native auto-page behavior distinct from PptxGenJS 4.0.1 defects', async () => {
    const document = PptxDocument.create();
    const source = document.addSlide();
    const following = document.addSlide();
    following.addText('FOLLOWING', { name: 'Native following sentinel' });
    const url = { url: 'https://auto-page.example?a=1&b=2', tooltip: 'Auto-page URL' };
    const internal = { slide: 2, tooltip: 'Existing following slide' };
    const rows = [
      ['Header A', 'Header B'],
      [
        { text: 'URL', options: { hyperlink: url } },
        { text: 'Internal', options: { hyperlink: internal } },
      ],
      [
        {
          text: [{
            runs: [
              { text: 'Rich', style: { bold: true, color: { kind: 'srgb', value: 'FF0000' } } },
              { text: 'body', style: { italic: true } },
            ],
          }],
        },
        'Rich value',
      ],
      ...Array.from({ length: 12 }, (_, index) => [
        `Body ${index}`,
        `Value ${index}`,
      ]),
    ] as const;
    const options = {
      autoPage: true,
      autoPageRepeatHeader: true,
      autoPageHeaderRows: 1,
      autoPageSlideStartY: inches(1.25),
      slideMargin: [inches(0.5), inches(0.5), inches(0.75), inches(0.5)] as const,
      x: inches(1),
      y: inches(4),
      columnWidths: [inches(4), inches(4)],
      rowHeights: [
        inches(0.25),
        ...Array.from({ length: 14 }, () => inches(0.375)),
      ],
    } as const;
    const callerSnapshot = JSON.stringify({ rows, options });
    source.addTable(rows, options);

    expect(JSON.stringify({ rows, options })).toBe(callerSnapshot);
    expect(Object.hasOwn(url, '_rId')).toBe(false);
    expect(Object.hasOwn(internal, '_rId')).toBe(false);
    expect(source.newAutoPagedSlides).toHaveLength(2);
    expect(document.slides).toEqual([
      source,
      ...source.newAutoPagedSlides,
      following,
    ]);
    expect(following.shapes.some(({ kind }) => kind === 'table')).toBe(false);
    const tables = [source, ...source.newAutoPagedSlides].map((slide) => slide.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    )!);
    expect(tables.map((table) => table.rows.map((row) => row.cells[0]!.text))).toEqual([
      ['Header A', 'URL'],
      ['Header A', 'Richbody', ...Array.from({ length: 8 }, (_, index) => `Body ${index}`)],
      ['Header A', ...Array.from({ length: 4 }, (_, index) => `Body ${index + 8}`)],
    ]);
    expect(tables[0]!.rows[1]!.cells.map(({ hyperlink }) => hyperlink)).toEqual([
      url,
      { slide: 4, tooltip: 'Existing following slide' },
    ]);

    for (const invalid of [
      { autoPage: 'true', rowHeights: [1] },
      { autoPage: true, rowHeights: [1], autoPageLineWeight: 1.001 },
      { autoPage: true, rowHeights: [1], autoPageCharWeight: '1' },
      { autoPage: true, rowHeights: [1], addHeaderToEach: true },
      { autoPage: true, rowHeights: [1], newSlideStartY: 1 },
    ]) {
      expect(() => source.addTable([['Invalid']], invalid as never)).toThrow();
    }

    const pptxgenjs = new PptxGenJS();
    const pptxgenjsSource = pptxgenjs.addSlide();
    const clamped: Record<string, unknown> = {
      x: 1,
      y: 1,
      w: 8,
      autoPage: true,
      autoPageLineWeight: 9,
    };
    pptxgenjsSource.addTable([['Clamped']], clamped);
    expect(clamped).toMatchObject({ autoPage: false, autoPageLineWeight: 1 });
    const coercible: Record<string, unknown> = {
      x: 1,
      y: 1,
      w: 8,
      autoPage: 'true',
    };
    expect(() => pptxgenjs.addSlide().addTable([['Coercible']], coercible))
      .not.toThrow();
    expect(coercible.autoPage).toBe(false);

    const overriddenCell = {
      text: 'Cell override',
      options: { autoPageCharWeight: -1 } as Record<string, unknown>,
    };
    const overridden = new PptxGenJS();
    overridden.addSlide().addTable([[overriddenCell]], {
      x: 1,
      y: 1,
      w: 8,
      autoPage: true,
      autoPageCharWeight: 1,
    });
    expect(overriddenCell.options.autoPageCharWeight).toBe(1);
  });

  it('locks the remaining addTable core declarations against PptxGenJS 4.0.1', async () => {
    const propertyId = (owner: string, property: string) =>
      `interface:${owner}@property:${property}`;
    const unionId = (owner: string, property: string, token: string) =>
      `union:${propertyId(owner, property)}#${token}`;
    const atomGroups = {
      supported: [
        ...['align', 'colspan', 'hyperlink', 'rowspan', 'transparency', 'valign']
          .map((property) => propertyId('TableCellProps', property)),
        ...[
          'align',
          'autoPage',
          'autoPageCharWeight',
          'autoPageHeaderRows',
          'autoPageLineWeight',
          'autoPageRepeatHeader',
          'valign',
        ].map((property) => propertyId('TableProps', property)),
      ],
      deliberate: [
        'method:Slide#addTable',
        ...['color', 'pt', 'type'].map((property) =>
          propertyId('BorderProps', property)),
        ...['dash', 'none', 'solid'].map((token) =>
          unionId('BorderProps', 'type', token)),
        ...['border', 'margin'].map((property) =>
          propertyId('TableCellProps', property)),
        unionId('TableCellProps', 'border', 'BorderProps'),
        unionId(
          'TableCellProps',
          'border',
          '[BorderProps,BorderProps,BorderProps,BorderProps]',
        ),
        ...[
          'autoPageSlideStartY',
          'border',
          'colW',
          'h',
          'margin',
          'objectName',
          'rowH',
          'verbose',
          'w',
          'x',
          'y',
        ].map((property) => propertyId('TableProps', property)),
        unionId('TableProps', 'border', 'BorderProps'),
        unionId(
          'TableProps',
          'border',
          '[BorderProps,BorderProps,BorderProps,BorderProps]',
        ),
        ...['number', 'number[]'].flatMap((token) => [
          unionId('TableProps', 'colW', token),
          unionId('TableProps', 'rowH', token),
        ]),
      ],
      deprecated: [propertyId('TableProps', 'newSlideStartY')],
      defect: [
        propertyId('TableCellProps', 'autoPageCharWeight'),
        propertyId('TableCellProps', 'autoPageLineWeight'),
        propertyId('TableProps', 'transparency'),
      ],
    } as const;
    expect(Object.fromEntries(Object.entries(atomGroups).map(([status, ids]) => [
      status,
      ids.length,
    ]))).toEqual({ supported: 13, deliberate: 28, deprecated: 1, defect: 3 });
    expect(new Set(Object.values(atomGroups).flat()).size).toBe(45);

    const generated = new PptxGenJS();
    generated.layout = 'LAYOUT_WIDE';
    const generatedSlide = generated.addSlide();
    const generatedLink: { url: string; tooltip: string; _rId?: number } = {
      url: 'https://add-table.example?a=1&b=2',
      tooltip: 'addTable family',
    };
    const tableOptions: Record<string, unknown> = {
      x: 0.5,
      y: 0.5,
      w: 6,
      h: 1.5,
      objectName: 'Audit & Table',
      colW: [1, 2, 3],
      rowH: [0.5, 0.5, 0.5],
      border: { type: 'dash', color: '4472C4', pt: 1.5 },
      margin: 0.1,
      align: 'center',
      valign: 'middle',
    };
    const returned = generatedSlide.addTable([
      [
        {
          text: 'Merge anchor',
          options: {
            colspan: 2,
            rowspan: 2,
            hyperlink: generatedLink,
            transparency: 25,
            color: '445566',
            margin: [1, 2, 3, 4],
            align: 'right',
            valign: 'bottom',
            border: [
              { type: 'none' },
              { type: 'dash', color: '00FF00', pt: 1 },
              { type: 'solid', color: '0000FF', pt: 2 },
              { type: 'solid', color: 'FF0000', pt: 3 },
            ],
          },
        },
        'Top right',
      ],
      ['Bottom right'],
      ['Footer 1', 'Footer 2', 'Footer 3'],
    ], tableOptions);
    generatedSlide.addTable([['Scalar A', 'Scalar B']], {
      x: 7,
      y: 0.5,
      w: 2.5,
      h: 0.5,
      objectName: 'Scalar table',
      colW: 1.25,
      rowH: 0.5,
      border: [
        { type: 'solid', color: 'FF0000', pt: 1 },
        { type: 'none' },
        { type: 'dash', color: '70AD47', pt: 2 },
        { type: 'solid', color: '4472C4', pt: 3 },
      ],
      align: 'left',
      valign: 'top',
    });

    expect(returned).toBe(generatedSlide);
    expect(tableOptions.objectName).toBe('Audit &amp; Table');
    const imported = await openPptxGenJSPublicOutput(generated);
    const [table, scalarTable] = imported.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(table).toBeInstanceOf(TableModel);
    expect(table!.name).toBe('Audit & Table');
    expect(table!.transform).toMatchObject({
      x: inches(0.5),
      y: inches(0.5),
      width: inches(6),
      height: inches(1.5),
    });
    expect(table!.columnWidths).toEqual([inches(1), inches(2), inches(3)]);
    expect(table!.rowHeights).toEqual(Array(3).fill(inches(0.5)));
    expect(table!.mergeRegions).toEqual([
      { rowIndex: 0, columnIndex: 0, rowspan: 2, colspan: 2 },
    ]);
    const anchor = table!.rows[0]!.cells[0]!;
    expect(anchor.hyperlink).toEqual({
      url: generatedLink.url,
      tooltip: generatedLink.tooltip,
    });
    expect(anchor.horizontalAlignment).toBe('right');
    expect(anchor.verticalAlignment).toBe('bottom');
    expect(anchor.richText[0]!.runs[0]!.style?.transparency).toBe(25);
    expect(anchor.margins).toBeDefined();
    expect(generatedLink._rId).toBeTypeOf('number');
    expect(table!.horizontalAlignment).toBeUndefined();
    expect(table!.verticalAlignment).toBeUndefined();
    expect(scalarTable!.name).toBe('Scalar table');
    expect(scalarTable!.columnWidths).toEqual(Array(2).fill(inches(1)));
    expect(scalarTable!.rowHeights).toEqual([inches(0.5)]);
    expect(scalarTable!.horizontalAlignment).toBe('left');
    expect(scalarTable!.verticalAlignment).toBe('top');
    const familyXml = slideXml(imported, 0);
    expect(familyXml).toContain('<a:alpha val="75000"/>');
    expect(familyXml).toContain('<a:prstDash val="sysDash"/>');
    expect(familyXml).toContain('<a:prstDash val="solid"/>');
    expect(familyXml).toContain('<a:noFill/>');
    expect(familyXml).toContain('<a:srgbClr val="00FF00"/>');
    expect(familyXml).toContain('<a:srgbClr val="0000FF"/>');
    expect(familyXml).toContain('<a:srgbClr val="FF0000"/>');

    const renderTransparency = async (
      tableTransparency: number | undefined,
      cellTransparency: number | undefined,
    ): Promise<string> => {
      const presentation = new PptxGenJS();
      const options: Record<string, unknown> = {
        x: 1,
        y: 1,
        w: 4,
        h: 1,
        color: '445566',
      };
      if (tableTransparency !== undefined) options.transparency = tableTransparency;
      const cellOptions: Record<string, unknown> = {};
      if (cellTransparency !== undefined) cellOptions.transparency = cellTransparency;
      presentation.addSlide().addTable(
        [[{ text: 'Transparency control', options: cellOptions }]],
        options,
      );
      return slideXml(await openPptxGenJSPublicOutput(presentation), 0);
    };
    const transparencyBaseline = await renderTransparency(undefined, undefined);
    const tableTransparency = await renderTransparency(50, undefined);
    const cellTransparency = await renderTransparency(undefined, 25);
    expect(tableTransparency).toBe(transparencyBaseline);
    expect(tableTransparency).not.toContain('<a:alpha val="50000"/>');
    expect(cellTransparency).toContain('<a:alpha val="75000"/>');

    const pagingRows = [
      ['Header'],
      ...Array.from(
        { length: 10 },
        (_, index) => [`Body ${index} ${'content '.repeat(12)}`],
      ),
    ];
    const createPaged = (
      startProperty: 'autoPageSlideStartY' | 'newSlideStartY',
      verbose = false,
    ) => {
      const presentation = new PptxGenJS();
      const source = presentation.addSlide();
      const options: Record<string, unknown> = {
        x: 1,
        y: 4,
        w: 5,
        autoPage: true,
        autoPageRepeatHeader: true,
        autoPageHeaderRows: 1,
        autoPageCharWeight: 0,
        autoPageLineWeight: 0,
        slideMargin: 0.5,
        fontSize: 18,
        margin: 0.05,
        verbose,
        [startProperty]: 1.25,
      };
      source.addTable(structuredClone(pagingRows), options);
      return { presentation, source };
    };
    const alias = createPaged('newSlideStartY');
    const canonical = createPaged('autoPageSlideStartY');
    const consoleLog = vi.spyOn(console, 'log').mockImplementation(() => undefined);
    let verbose: ReturnType<typeof createPaged> | undefined;
    let verboseMessages = '';
    try {
      verbose = createPaged('autoPageSlideStartY', true);
    } finally {
      verboseMessages = consoleLog.mock.calls.flat().map(String).join('\n');
      consoleLog.mockRestore();
    }
    expect(verboseMessages).toContain('[[VERBOSE MODE]]');
    expect(verboseMessages).toContain('NEW SLIDE CREATED');
    expect(alias.source.newAutoPagedSlides.length).toBeGreaterThan(0);
    expect(alias.source.newAutoPagedSlides).toHaveLength(
      canonical.source.newAutoPagedSlides.length,
    );
    expect(verbose!.source.newAutoPagedSlides).toHaveLength(
      canonical.source.newAutoPagedSlides.length,
    );
    const [aliasDocument, canonicalDocument, verboseDocument] = await Promise.all([
      openPptxGenJSPublicOutput(alias.presentation),
      openPptxGenJSPublicOutput(canonical.presentation),
      openPptxGenJSPublicOutput(verbose!.presentation),
    ]);
    const tableState = (document: PptxDocument) => document.slides.map((slide) => {
      const candidate = slide.shapes.find(
        (shape): shape is TableModel => shape instanceof TableModel,
      )!;
      return {
        y: candidate.transform.y,
        rows: candidate.rows.map((row) => row.cells[0]!.text),
      };
    });
    expect(tableState(aliasDocument)).toEqual(tableState(canonicalDocument));
    expect(tableState(verboseDocument)).toEqual(tableState(canonicalDocument));
    expect(tableState(aliasDocument).slice(1).every(({ y }) => y === inches(1.25)))
      .toBe(true);

    const overriddenCell = {
      text: 'Cell char weight',
      options: { margin: 0, autoPageCharWeight: -1 } as Record<string, unknown>,
    };
    const charWeight = new PptxGenJS();
    charWeight.addSlide().addTable([[overriddenCell], ['Second cell weight row']], {
      x: 1,
      y: 4,
      w: 2,
      autoPage: true,
      autoPageCharWeight: 1,
      margin: 0,
    });
    expect(overriddenCell.options.autoPageCharWeight).toBe(1);

    const measurementRows = Array.from({ length: 8 }, (_, index) => [{
      text: `${index} ${'AAAAAAAAAA '.repeat(8)}`,
      options: { margin: 0 } as Record<string, unknown>,
    }]);
    const pageCount = (
      cellLineWeight: number,
      tableLineWeight: number,
    ): { readonly pages: number; readonly local: number } => {
      const presentation = new PptxGenJS();
      presentation.defineLayout({ name: 'ADD-TABLE-WEIGHT', width: 10, height: 5.625 });
      presentation.layout = 'ADD-TABLE-WEIGHT';
      const source = presentation.addSlide();
      const rows = structuredClone(measurementRows);
      rows.forEach((row) => {
        row[0]!.options.autoPageLineWeight = cellLineWeight;
      });
      source.addTable(rows, {
        x: 1,
        y: 4,
        w: 2,
        autoPage: true,
        autoPageSlideStartY: 1,
        slideMargin: 0.5,
        fontSize: 12,
        margin: 0,
        autoPageLineWeight: tableLineWeight,
      });
      return {
        pages: source.newAutoPagedSlides.length,
        local: rows[0]![0]!.options.autoPageLineWeight as number,
      };
    };
    const cellMinus = pageCount(-1, 0);
    const cellPlus = pageCount(1, 0);
    expect(cellMinus.pages).toBe(cellPlus.pages);
    expect([cellMinus.local, cellPlus.local]).toEqual([-1, 1]);
    expect(pageCount(-1, -1).pages).toBeLessThan(pageCount(1, 1).pages);
  }, 60_000);

  it('matches legal PptxGenJS automatic table boundaries at all public weights', async () => {
    const weights = [-1, 0, 1] as const;
    const bodyTexts = Array.from(
      { length: 4 },
      (_, index) => `${index} ${Array.from({ length: 5 }, () => 'A'.repeat(10)).join(' ')}`,
    );
    const headerText = `H ${Array.from({ length: 5 }, () => 'A'.repeat(10)).join(' ')}`;

    for (const weight of weights) {
      const expectedLines = weight === -1 ? 5 : weight === 0 ? 3 : 2;
      const lineHeight = 12 * (1.67 + weight) / 100;
      const pageCapacity = lineHeight * expectedLines + 0.001;
      const pageStartY = 5.125 - pageCapacity;
      const generated = new PptxGenJS();
      generated.defineLayout({ name: 'NATIVE-SIZE', width: 10, height: 5.625 });
      generated.layout = 'NATIVE-SIZE';
      const generatedSlide = generated.addSlide();
      const generatedCellOptions = Array.from(
        { length: bodyTexts.length + 1 },
        () => ({ margin: 0 } as Record<string, unknown>),
      );
      const generatedRows: (string | PptxGenJSTableCell)[][] = [
        [{ text: headerText, options: generatedCellOptions[0]! }],
        ...bodyTexts.map((text, index) => [{
          text,
          options: generatedCellOptions[index + 1]!,
        }]),
      ];
      const generatedOptions: Record<string, unknown> = {
        x: 1,
        y: pageStartY,
        w: 2,
        autoPage: true,
        autoPageCharWeight: weight,
        autoPageLineWeight: weight,
        autoPageSlideStartY: pageStartY,
        slideMargin: 0.5,
        fontSize: 12,
        margin: 0,
      };
      generatedSlide.addTable(generatedRows, generatedOptions);
      expect(generatedSlide.newAutoPagedSlides.length, `PptxGenJS weight ${weight}`)
        .toBeGreaterThan(0);
      expect(generatedOptions, `PptxGenJS weight ${weight}`).toMatchObject({
        autoPage: false,
        autoPageCharWeight: weight,
        autoPageLineWeight: weight,
      });
      expect(generatedCellOptions.every(({ autoPageCharWeight }) =>
        autoPageCharWeight === (weight === 0 ? null : weight))).toBe(true);

      const imported = await importPptxGenJS(generated);
      const generatedHeights: readonly (readonly number[])[] = imported.slides.map((slide) => {
        const table = slide.shapes.find(
          (shape): shape is TableModel => shape instanceof TableModel,
        )!;
        return table.rowHeights!;
      });
      const generatedPages = imported.slides.map((slide) => {
        const table = slide.shapes.find(
          (shape): shape is TableModel => shape instanceof TableModel,
        )!;
        return table.rows.map((row) => row.cells[0]!.text);
      });

      const native = PptxDocument.create({
        slideSize: { width: inches(10), height: inches(5.625) },
      });
      const nativeSlide = native.addSlide();
      const nativeRows: readonly (readonly AddTableCellInput[])[] = [
        [{ text: headerText, options: { margin: 0 } }],
        ...bodyTexts.map((text) => [{ text, options: { margin: 0 } }]),
      ];
      const nativeSnapshot = structuredClone(nativeRows);
      nativeSlide.addTable(nativeRows, {
        autoPage: true,
        autoPageCharWeight: weight,
        autoPageLineWeight: weight,
        autoPageSlideStartY: inches(pageStartY),
        slideMargin: inches(0.5),
        x: inches(1),
        y: inches(pageStartY),
        columnWidths: [inches(2)],
        fontSize: 12,
        margin: 0,
      });
      expect(nativeRows, `native detachment weight ${weight}`).toEqual(nativeSnapshot);
      const nativePages = [nativeSlide, ...nativeSlide.newAutoPagedSlides].map((slide) => {
        const table = slide.shapes.find(
          (shape): shape is TableModel => shape instanceof TableModel,
        )!;
        return table.rows.map((row) => row.cells[0]!.text);
      });
      expect(nativePages, `weight ${weight}`).toEqual(generatedPages);
      const nativeHeights = [nativeSlide, ...nativeSlide.newAutoPagedSlides].map((slide) => {
        const table = slide.shapes.find(
          (shape): shape is TableModel => shape instanceof TableModel,
        )!;
        return table.rowHeights!;
      });
      expect(generatedHeights.flat().every((height) => height === 0), `PptxGenJS weight ${weight}`)
        .toBe(true);
      expect(nativeHeights.flat(), `native row heights weight ${weight}`).toEqual(
        Array.from(
          { length: bodyTexts.length + 1 },
          () => expectedLines * inches(lineHeight),
        ),
      );
      expect(nativePages.flat(), `weight ${weight}`).toEqual([headerText, ...bodyTexts]);
    }
  }, 60_000);

  it('imports and edits legal PptxGenJS plain table-cell hyperlinks', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const source = generated.addSlide();
    generated.addSlide();
    const generatedUrl: {
      url: string;
      tooltip?: string;
      _rId?: number;
    } = {
      url: 'https://example.com?a=1&b=2',
      tooltip: 'Visit & learn',
    };
    const generatedInternal: {
      slide: number;
      tooltip?: string;
      _rId?: number;
    } = { slide: 2 };
    source.addTable([[
      { text: 'URL', options: { hyperlink: generatedUrl } },
      { text: 'Slide', options: { hyperlink: generatedInternal } },
      { text: 'Plain', options: {} },
    ]], { x: 1, y: 1, w: 8, h: 1 });

    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes[0] as TableModel;
    expect(importedTable.rows[0]!.cells.map(({ hyperlink }) => hyperlink)).toEqual([
      {
        url: 'https://example.com?a=1&b=2',
        tooltip: 'Visit & learn',
      },
      { slide: 2, tooltip: '' },
      undefined,
    ]);
    expect(generatedUrl._rId).toBeTypeOf('number');
    expect(generatedInternal._rId).toBeTypeOf('number');
    const importedRelationships = imported.slides[0]!.relationships;
    expect(importedRelationships.find(({ type }) => type.endsWith('/hyperlink')))
      .toMatchObject({
        target: 'https://example.com?a=1&b=2',
        targetMode: 'External',
      });
    expect(importedRelationships.find(({ resolvedTarget }) =>
      resolvedTarget === imported.slides[1]!.partUri)).toBeDefined();
    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    expect(importedXml.match(/<a:hlinkClick\b/g)).toHaveLength(2);
    expect(importedXml.match(/\bu="sng"/g)).toHaveLength(2);
    expect(importedXml).toContain('invalidUrl="" action=""');
    expect(importedXml).toContain('action="ppaction://hlinksldjump" tooltip=""');

    const callerUrlRelationshipId = generatedUrl._rId;
    const callerInternalRelationshipId = generatedInternal._rId;
    const externalRelationshipId = importedRelationships.find(
      ({ type }) => type.endsWith('/hyperlink'),
    )!.id;
    const internalRelationshipId = importedRelationships.find(
      ({ resolvedTarget }) => resolvedTarget === imported.slides[1]!.partUri,
    )!.id;
    importedTable.setCellHyperlink(0, 0, {
      url: 'https://adapter-edited.example?a=1&b=2',
      tooltip: '',
    });
    importedTable.setCellHyperlink(0, 1, { slide: 2, tooltip: 'Adapter target' });
    importedTable.setCellHyperlink(0, 2, { url: 'https://temporary.example' });
    importedTable.setCellHyperlink(0, 2, undefined);
    expect(importedTable.rows[0]!.cells.map(({ hyperlink }) => hyperlink)).toEqual([
      { url: 'https://adapter-edited.example?a=1&b=2', tooltip: '' },
      { slide: 2, tooltip: 'Adapter target' },
      undefined,
    ]);
    expect(generatedUrl._rId).toBe(callerUrlRelationshipId);
    expect(generatedInternal._rId).toBe(callerInternalRelationshipId);
    expect(imported.slides[0]!.relationships.find(({ id }) => id === externalRelationshipId))
      .toMatchObject({
        target: 'https://adapter-edited.example?a=1&b=2',
        targetMode: 'External',
      });
    expect(imported.slides[0]!.relationships.find(({ id }) => id === internalRelationshipId))
      .toMatchObject({ resolvedTarget: imported.slides[1]!.partUri });
    const editedImportedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    const editedExternalClick = editedImportedXml.match(
      new RegExp(`<a:hlinkClick\\b[^>]*r:id="${externalRelationshipId}"[^>]*>`),
    )?.[0];
    expect(editedExternalClick).toContain('invalidUrl=""');
    expect(editedExternalClick).toContain('tgtFrame=""');
    expect(editedExternalClick).toContain('history="1"');
    expect(editedExternalClick).toContain('highlightClick="0"');
    expect(editedExternalClick).toContain('endSnd="0"');
    expect(editedExternalClick).toContain('tooltip=""');
    expect(editedExternalClick).not.toContain('action=""');

    const native = PptxDocument.create();
    const nativeSource = native.addSlide();
    native.addSlide();
    const nativeUrl = {
      url: 'https://example.com?a=1&b=2',
      tooltip: 'Visit & learn',
    };
    const nativeTable = nativeSource.addTable([[
      { text: 'URL', options: { hyperlink: nativeUrl } },
      { text: 'Slide', options: { hyperlink: { slide: 2 } } },
      'Plain',
    ]], {
      x: inches(1),
      y: inches(1),
      width: inches(8),
      height: inches(1),
    });
    expect(nativeTable.rows[0]!.cells.map(({ hyperlink }) => hyperlink)).toEqual([
      {
        url: 'https://example.com?a=1&b=2',
        tooltip: 'Visit & learn',
      },
      { slide: 2 },
      undefined,
    ]);
    expect(Object.hasOwn(nativeUrl, '_rId')).toBe(false);
    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(nativeSource.partUri).bytes,
    );
    expect(nativeXml.match(/<a:hlinkClick\b/g)).toHaveLength(2);
    expect(nativeXml.match(/\bu="sng"/g)).toHaveLength(2);
    expect(nativeXml).not.toContain('invalidUrl=');
    expect(nativeXml).toMatch(
      /<a:hlinkClick r:id="rId\d+" action="ppaction:\/\/hlinksldjump"\/>/,
    );

    const reopenedImported = await PptxDocument.open(await imported.write());
    const reopenedNative = await PptxDocument.open(await native.write());
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel)
      .rows[0]!.cells.map(({ hyperlink }) => hyperlink)).toEqual([
      {
        url: 'https://adapter-edited.example?a=1&b=2',
        tooltip: '',
      },
      { slide: 2, tooltip: 'Adapter target' },
      undefined,
    ]);
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel)
      .rows[0]!.cells.map(({ hyperlink }) => hyperlink)).toEqual([
      {
        url: 'https://example.com?a=1&b=2',
        tooltip: 'Visit & learn',
      },
      { slide: 2 },
      undefined,
    ]);
  }, 20_000);

  it('imports and edits legal PptxGenJS rich table-cell text semantically', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    generated.addSlide();
    const outer: { url: string; tooltip: string; _rId?: number } = {
      url: 'https://outer-rich.example',
      tooltip: '',
    };
    const local: { slide: number; tooltip: string; _rId?: number } = {
      slide: 2,
      tooltip: 'Local target',
    };
    generatedSlide.addTable([[
      {
        text: [
          { text: 'Outer', options: { bold: true } },
          {
            text: ' local',
            options: { italic: true, hyperlink: local, breakLine: true },
          },
          { text: 'Second', options: { color: '336699', underline: true } },
        ],
        options: { hyperlink: outer },
      },
    ]], { x: 1, y: 1, w: 6, h: 1.5 });

    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes[0] as TableModel;
    const importedCell = importedTable.rows[0]!.cells[0]!;
    expect(importedCell.text).toBe('Outer local\nSecond');
    expect(importedCell.richText).toHaveLength(2);
    expect(importedCell.richText.map(({ runs }) => runs.map(({ text }) => text)))
      .toEqual([['Outer', ' local'], ['Second']]);
    expect(importedCell.richText[0]!.runs[0]!.style).toMatchObject({
      bold: true,
      hyperlink: { url: 'https://outer-rich.example', tooltip: '' },
    });
    expect(importedCell.richText[0]!.runs[1]!.style).toMatchObject({
      italic: true,
      hyperlink: { slide: 2, tooltip: 'Local target' },
    });
    expect(importedCell.richText[1]!.runs[0]!.style).toMatchObject({
      color: { kind: 'srgb', value: '336699' },
      hyperlink: { url: 'https://outer-rich.example', tooltip: '' },
    });
    expect(outer._rId).toBeTypeOf('number');
    expect(local._rId).toBeTypeOf('number');

    importedTable.setCellRichText(0, 0, [{
      align: 'right',
      runs: [
        {
          text: 'Edited',
          style: { bold: true, hyperlink: {
            url: 'https://adapter-rich-edited.example',
            tooltip: '',
          } },
        },
        {
          text: ' target',
          style: { italic: true, hyperlink: { slide: 2, tooltip: 'Edited target' } },
        },
      ],
    }]);
    expect(importedTable.rows[0]!.cells[0]!.text).toBe('Edited target');
    expect(importedTable.rows[0]!.cells[0]!.richText[0]!.runs.map(
      ({ style }) => style?.hyperlink,
    )).toEqual([
      { url: 'https://adapter-rich-edited.example', tooltip: '' },
      { slide: 2, tooltip: 'Edited target' },
    ]);

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    native.addSlide();
    const nativeOuter = { url: 'https://outer-rich.example', tooltip: '' };
    const nativeTable = nativeSlide.addTable([[
      {
        text: [
          { runs: [
            { text: 'Outer', style: { bold: true } },
            { text: ' local', style: {
              italic: true,
              hyperlink: { slide: 2, tooltip: 'Local target' },
            } },
          ] },
          { runs: [{
            text: 'Second',
            style: {
              color: { kind: 'srgb', value: '336699' },
              underline: true,
            },
          }] },
        ],
        options: { hyperlink: nativeOuter },
      },
    ]], {
      x: inches(1),
      y: inches(1),
      width: inches(6),
      height: inches(1.5),
    });
    expect(nativeTable.rows[0]!.cells[0]!.text).toBe('Outer local\nSecond');
    expect(nativeTable.rows[0]!.cells[0]!.richText.map(({ runs }) =>
      runs.map(({ text }) => text))).toEqual(
      importedCell.richText.map(({ runs }) => runs.map(({ text }) => text)),
    );
    expect(Object.hasOwn(nativeOuter, '_rId')).toBe(false);

    const reopenedImported = await PptxDocument.open(await imported.write());
    const reopenedTable = reopenedImported.slides[0]!.shapes[0] as TableModel;
    expect(reopenedTable.rows[0]!.cells[0]!.text).toBe('Edited target');
    expect(reopenedTable.rows[0]!.cells[0]!.richText[0]!.runs[1]!.style?.hyperlink)
      .toEqual({ slide: 2, tooltip: 'Edited target' });
  }, 20_000);

  it('imports PptxGenJS table text defaults and locks native differences', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const generatedSlide = generated.addSlide();
    const tableOptions = {
      x: 0.5,
      y: 0.5,
      w: 12,
      h: 1.5,
      fontFace: 'Aptos',
      fontSize: 18.25,
      bold: true,
      color: '4472C4',
      paraSpaceBefore: 6,
      paraSpaceAfter: 8,
      lineSpacingMultiple: 1.5,
    };
    const falseCellOptions = {
      fontFace: 'Courier New',
      fontSize: 10,
      bold: false,
      color: '00AA00',
      paraSpaceBefore: 3,
      paraSpaceAfter: 9,
      lineSpacing: 22,
    };
    const richCellOptions = { lineSpacingMultiple: 1.25 };
    const falseRunOptions = {
      fontFace: 'Arial',
      fontSize: 8,
      bold: false,
      color: 'FF0000',
    };
    const defaultHyperlink: { url: string; tooltip: string; _rId?: number } = {
      url: 'https://default-table-link.example',
      tooltip: '',
    };
    const localHyperlink: { url: string; tooltip: string; _rId?: number } = {
      url: 'https://local-table-link.example',
      tooltip: 'Local',
    };
    generatedSlide.addTable([[
      { text: 'Plain', options: {} },
      { text: 'Cell false', options: falseCellOptions },
      {
        text: [
          { text: 'Rich', options: {} },
          { text: ' local false', options: falseRunOptions },
        ],
        options: richCellOptions,
      },
      { text: '', options: {} },
      { text: 'Default link', options: { hyperlink: defaultHyperlink } },
      {
        text: [{ text: 'Local link', options: { hyperlink: localHyperlink } }],
        options: {},
      },
    ]], tableOptions);

    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes[0] as TableModel;
    const importedCells = importedTable.rows[0]!.cells;
    expect(importedCells[0]!.richText[0]!.runs[0]!.style).toMatchObject({
      fontFamily: 'Aptos',
      fontSize: 18.25,
      bold: true,
      color: { kind: 'srgb', value: '4472C4' },
    });
    expect(importedCells[0]!.richText[0]!.spacing).toBeUndefined();
    expect(importedCells[1]!.richText[0]).toMatchObject({
      spacing: {
        before: 3,
        after: 9,
        line: { kind: 'exact', points: 22 },
      },
      runs: [{
        style: {
          fontFamily: 'Courier New',
          fontSize: 10,
          bold: true,
          color: { kind: 'srgb', value: '00AA00' },
        },
      }],
    });
    expect(importedCells[2]!.richText[0]!.spacing).toEqual({
      line: { kind: 'multiple', factor: 1.25 },
    });
    expect(importedCells[2]!.richText[0]!.runs).toMatchObject([
      {
        text: 'Rich',
        style: {
          fontFamily: 'Aptos',
          fontSize: 18.25,
          bold: true,
          color: { kind: 'srgb', value: '4472C4' },
        },
      },
      {
        text: ' local false',
        style: {
          fontFamily: 'Arial',
          fontSize: 8,
          bold: true,
          color: { kind: 'srgb', value: 'FF0000' },
        },
      },
    ]);
    expect(importedCells[3]!.richText).toMatchObject([{ runs: [] }]);
    for (const [index, hyperlink] of [
      [4, { url: 'https://default-table-link.example', tooltip: '' }],
      [5, { url: 'https://local-table-link.example', tooltip: 'Local' }],
    ] as const) {
      expect(importedCells[index]!.richText[0]!.runs[0]!.style).toMatchObject({
        fontFamily: 'Aptos',
        fontSize: 18.25,
        bold: true,
        hyperlink,
      });
    }
    expect(importedCells[4]!.richText[0]!.runs[0]!.style?.color).toEqual({
      kind: 'srgb',
      value: '4472C4',
    });
    expect(importedCells[5]!.richText[0]!.runs[0]!.style?.color).toBeUndefined();
    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    const importedCellXml = [...importedXml.matchAll(/<a:tc>[\s\S]*?<\/a:tc>/g)]
      .map(([cell]) => cell);
    expect(importedCellXml[3]).toContain(
      '<a:endParaRPr lang="en-US" sz="1825" dirty="0">',
    );
    expect(importedCellXml[3]).toContain('typeface="Aptos"');

    expect(tableOptions).toMatchObject({
      x: inches(0.5),
      y: inches(0.5),
      w: inches(12),
      h: inches(1.5),
      margin: [0.05, 0.1, 0.05, 0.1],
    });
    expect(falseCellOptions.bold).toBe(true);
    expect(falseRunOptions.bold).toBe(true);
    expect(Object.hasOwn(falseCellOptions, '_lineIdx')).toBe(true);
    expect(Object.hasOwn(falseRunOptions, '_lineIdx')).toBe(true);
    expect(defaultHyperlink._rId).toBeTypeOf('number');
    expect(localHyperlink._rId).toBeTypeOf('number');

    const native = PptxDocument.create();
    const nativeSlide = native.addSlide();
    const nativeTableOptions = {
      fontFamily: 'Aptos',
      fontSize: 18.25,
      bold: true,
      color: { kind: 'srgb' as const, value: '4472C4' },
    };
    const nativeFalseCellOptions = {
      fontFamily: 'Courier New',
      fontSize: 10,
      bold: false,
      color: { kind: 'srgb' as const, value: '00AA00' },
      spacing: {
        before: 3,
        after: 9,
        line: { kind: 'exact' as const, points: 22 },
      },
    };
    const nativeFalseRunStyle = {
      fontFamily: 'Arial',
      fontSize: 8,
      bold: false,
      color: { kind: 'srgb' as const, value: 'FF0000' },
    };
    const nativeInputsBefore = JSON.stringify({
      nativeTableOptions,
      nativeFalseCellOptions,
      nativeFalseRunStyle,
    });
    const nativeTable = nativeSlide.addTable([[
      'Plain',
      { text: 'Cell false', options: nativeFalseCellOptions },
      {
        text: [{ runs: [
          { text: 'Rich' },
          { text: ' local false', style: nativeFalseRunStyle },
        ] }],
        options: { spacing: { line: { kind: 'multiple', factor: 1.25 } } },
      },
      { text: [{ runs: [] }] },
      { text: 'Default link', options: {
        hyperlink: { url: 'https://default-table-link.example', tooltip: '' },
      } },
      { text: [{ runs: [{
        text: 'Local link',
        style: { hyperlink: {
          url: 'https://local-table-link.example',
          tooltip: 'Local',
        } },
      }] }] },
    ]], nativeTableOptions);
    const nativeCells = nativeTable.rows[0]!.cells;
    expect(nativeCells[0]!.richText).toEqual(importedCells[0]!.richText);
    expect(nativeCells[1]!.richText[0]!.spacing).toEqual(
      importedCells[1]!.richText[0]!.spacing,
    );
    expect(nativeCells[1]!.richText[0]!.runs[0]!.style).toMatchObject({
      fontFamily: 'Courier New',
      fontSize: 10,
      bold: false,
      color: { kind: 'srgb', value: '00AA00' },
    });
    expect(nativeCells[2]!.richText[0]!.runs[0]).toEqual(
      importedCells[2]!.richText[0]!.runs[0],
    );
    expect(nativeCells[2]!.richText[0]!.runs[1]!.style).toMatchObject({
      fontFamily: 'Arial',
      fontSize: 8,
      bold: false,
      color: { kind: 'srgb', value: 'FF0000' },
    });
    expect(nativeCells[3]!.richText).toEqual(importedCells[3]!.richText);
    expect(nativeCells.slice(4).map(({ richText }) => richText))
      .toEqual(importedCells.slice(4).map(({ richText }) => richText));
    expect(JSON.stringify({
      nativeTableOptions,
      nativeFalseCellOptions,
      nativeFalseRunStyle,
    })).toBe(nativeInputsBefore);

    const nativeExtension = PptxDocument.create();
    const extensionTable = nativeExtension.addSlide().addTable([[
      { text: 'Native false', options: { bold: false } },
    ]], {
      bold: true,
      spacing: {
        before: 6,
        after: 8,
        line: { kind: 'multiple', factor: 1.5 },
      },
    });
    expect(extensionTable.rows[0]!.cells[0]!.richText[0]!.runs[0]!.style?.bold)
      .toBe(false);
    expect(extensionTable.rows[0]!.cells[0]!.richText[0]!.spacing).toEqual({
      before: 6,
      after: 8,
      line: { kind: 'multiple', factor: 1.5 },
    });

    importedTable.setCellText(0, 0, 'Imported edited');
    nativeTable.setCellText(0, 0, 'Native edited');
    importedTable.setCellRichText(0, 2, [{ runs: [{
      text: 'Imported rich edit',
      style: { bold: false, color: { kind: 'srgb', value: '112233' } },
    }] }]);
    nativeTable.setCellRichText(0, 2, [{ runs: [{
      text: 'Native rich edit',
      style: { bold: false, color: { kind: 'srgb', value: '112233' } },
    }] }]);
    const reopenedImported = await PptxDocument.open(await imported.write());
    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImportedTable = reopenedImported.slides[0]!.shapes[0] as TableModel;
    const reopenedNativeTable = reopenedNative.slides[0]!.shapes[0] as TableModel;
    expect(reopenedImportedTable.rows[0]!.cells[0]!.richText[0]!.runs[0]).toMatchObject({
      text: 'Imported edited',
      style: { fontFamily: 'Aptos', fontSize: 18.25, bold: true },
    });
    expect(reopenedNativeTable.rows[0]!.cells[0]!.richText[0]!.runs[0]).toMatchObject({
      text: 'Native edited',
      style: { fontFamily: 'Aptos', fontSize: 18.25, bold: true },
    });
    expect(reopenedImportedTable.rows[0]!.cells[2]!.richText[0]!.runs[0]).toMatchObject({
      text: 'Imported rich edit',
      style: { bold: false, color: { kind: 'srgb', value: '112233' } },
    });
    expect(reopenedNativeTable.rows[0]!.cells[2]!.richText[0]!.runs[0]).toMatchObject({
      text: 'Native rich edit',
      style: { bold: false, color: { kind: 'srgb', value: '112233' } },
    });
  }, 20_000);

  it('repairs a PptxGenJS transform and column-grid mismatch through the public model', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable(
      [['A', 'B', 'C'].map((text) => ({ text, options: {} }))],
      { x: 1, y: 1, w: 5, h: 1, colW: [1, 2, 3] },
    );

    const imported = await importPptxGenJS(generated);
    const table = imported.slides[0]!.shapes[0] as TableModel;
    expect(table).toBeInstanceOf(TableModel);
    expect(table.columnWidths).toEqual([
      inches(1),
      inches(2),
      inches(3),
    ]);
    expect(table.transform.width).toBe(inches(5));
    const before = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    const gridBefore = [...before.matchAll(/<a:gridCol w="(\d+)"\/>/g)]
      .map((match) => match[1]);

    table.setColumnWidths(table.columnWidths!);

    expect(table.transform.width).toBe(inches(6));
    const after = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    expect([...after.matchAll(/<a:gridCol w="(\d+)"\/>/g)]
      .map((match) => match[1])).toEqual(gridBefore);
    expect(after).toContain('<a:ext cx="5486400" cy="914400"/>');

    const reopened = await PptxDocument.open(await imported.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    expect(reopenedTable.columnWidths).toEqual([
      inches(1),
      inches(2),
      inches(3),
    ]);
    expect(reopenedTable.transform.width).toBe(inches(6));
  });

  it('repairs PptxGenJS scalar column-width floor while preserving public intent', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const rows = [['A', 'B', 'C']] as const;
    generated.addSlide().addTable(
      rows.map((row) => row.map((text) => ({ text, options: {} }))),
      { x: 1, y: 1, h: 1, colW: 1.25 },
    );
    const imported = await importPptxGenJS(generated);

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeTable = native.addSlide().addTable(rows, {
      x: inches(1),
      y: inches(1),
      height: inches(1),
      columnWidths: inches(1.25),
    });

    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(native.slides[0]!.partUri).bytes,
    );
    const readColumnWidths = (xml: string): number[] =>
      [...xml.matchAll(/<a:gridCol w="(\d+)"\/>/g)].map((match) => Number(match[1]));
    expect(readColumnWidths(importedXml)).toEqual([
      inches(1),
      inches(1),
      inches(1),
    ]);
    expect(readColumnWidths(nativeXml)).toEqual([
      inches(1.25),
      inches(1.25),
      inches(1.25),
    ]);
    expect((imported.slides[0]!.shapes[0] as TableModel).transform.width).toBe(inches(3));
    expect(nativeTable.transform.width).toBe(inches(3.75));
    expect(nativeXml).toContain('<a:ext cx="3429000" cy="914400"/>');
  });

  it('repairs an explicit PptxGenJS row-height mismatch and preserves transform height for automatic rows', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable(
      [
        ['A', 'B'].map((text) => ({ text, options: {} })),
        ['C', 'D'].map((text) => ({ text, options: {} })),
      ],
      { x: 1, y: 1, w: 5, h: 1, rowH: [0.5, 1.5] },
    );

    const imported = await importPptxGenJS(generated);
    const table = imported.slides[0]!.shapes[0] as TableModel;
    expect(table).toBeInstanceOf(TableModel);
    expect(table.rowHeights).toEqual([inches(0.5), inches(1.5)]);
    expect(table.transform.height).toBe(inches(1));
    const before = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    const rowsBefore = [...before.matchAll(/<a:tr h="(\d+)">/g)]
      .map((match) => match[1]);

    table.setRowHeights(table.rowHeights!);

    expect(table.transform.height).toBe(inches(2));
    let after = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    expect([...after.matchAll(/<a:tr h="(\d+)">/g)]
      .map((match) => match[1])).toEqual(rowsBefore);
    expect(after).toContain('cy="1828800"');

    table.setRowHeights([0, inches(1.5)]);
    expect(table.rowHeights).toEqual([0, inches(1.5)]);
    expect(table.transform.height).toBe(inches(2));
    after = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    expect(after).toContain('cy="1828800"');

    const reopened = await PptxDocument.open(await imported.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    expect(reopenedTable.rowHeights).toEqual([0, inches(1.5)]);
    expect(reopenedTable.transform.height).toBe(inches(2));
  });

  it('repairs PptxGenJS omitted-height mismatch for explicit row heights', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const rows = [['A'], ['B'], ['C']] as const;
    generated.addSlide().addTable(
      rows.map((row) => row.map((text) => ({ text, options: {} }))),
      { x: 1, y: 1, rowH: [0.5, 1, 1.5] },
    );
    const imported = await importPptxGenJS(generated);

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeTable = native.addSlide().addTable(rows, {
      x: inches(1),
      y: inches(1),
      rowHeights: [inches(0.5), inches(1), inches(1.5)],
    });

    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(native.slides[0]!.partUri).bytes,
    );
    const readRowHeights = (xml: string): number[] =>
      [...xml.matchAll(/<a:tr h="(\d+)">/g)].map((match) => Number(match[1]));
    const expectedHeights = [inches(0.5), inches(1), inches(1.5)];
    expect(readRowHeights(importedXml)).toEqual(expectedHeights);
    expect(readRowHeights(nativeXml)).toEqual(expectedHeights);
    expect((imported.slides[0]!.shapes[0] as TableModel).transform.height).toBe(inches(1));
    expect(nativeTable.transform.height).toBe(inches(3));
    expect(nativeXml).toContain('<a:ext cx="914400" cy="2743200"/>');
  });

  it('projects and normalizes PptxGenJS table-level border output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[{ text: 'Uniform A', options: {} }, { text: 'Uniform B', options: {} }]],
      {
        x: 0.5,
        y: 0.5,
        w: 4,
        h: 1,
        border: { type: 'dash', color: '4472C4', pt: 1.5 },
      },
    );
    slide.addTable(
      [[{ text: 'Omitted A', options: {} }, { text: 'Omitted B', options: {} }]],
      { x: 0.5, y: 2, w: 4, h: 1 },
    );
    slide.addTable(
      [[
        { text: 'Inherited', options: {} },
        { text: 'Override', options: { border: { type: 'none' } } },
      ]],
      {
        x: 0.5,
        y: 3.5,
        w: 4,
        h: 1,
        border: { type: 'solid', color: '0000FF', pt: 1 },
      },
    );
    slide.addTable(
      [[{ text: 'TRBL', options: {} }]],
      {
        x: 0.5,
        y: 5,
        w: 4,
        h: 1,
        border: [
          { type: 'solid', color: 'FF0000', pt: 1 },
          { type: 'none' },
          { type: 'dash', color: '70AD47', pt: 2 },
          { type: 'solid', color: '4472C4', pt: 3 },
        ],
      },
    );

    const imported = await importPptxGenJS(generated);
    const tables = imported.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    const four = <T>(border: T) => ({
      top: border,
      right: border,
      bottom: border,
      left: border,
    });
    const none = { kind: 'none' as const };
    const uniformLine = {
      kind: 'line' as const,
      color: { kind: 'srgb' as const, value: '4472C4' },
      width: 1.5,
      style: 'dash' as const,
    };
    const inheritedLine = {
      kind: 'line' as const,
      color: { kind: 'srgb' as const, value: '0000FF' },
      width: 1,
      style: 'solid' as const,
    };
    expect(tables).toHaveLength(4);
    expect(tables[0]!.borders).toEqual(four(uniformLine));
    expect(tables[0]!.rows[0]!.cells.map(({ borders }) => borders))
      .toEqual(Array(2).fill(four(uniformLine)));
    expect(tables[1]!.borders).toEqual(four(none));
    expect(tables[1]!.rows[0]!.cells.map(({ borders }) => borders))
      .toEqual(Array(2).fill(four(none)));
    expect(tables[2]!.borders).toBeUndefined();
    expect(tables[2]!.rows[0]!.cells.map(({ borders }) => borders)).toEqual([
      four(inheritedLine),
      four(none),
    ]);
    expect(tables[3]!.borders).toEqual({
      top: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        width: 1,
        style: 'solid',
      },
      right: none,
      bottom: {
        kind: 'line',
        color: { kind: 'srgb', value: '70AD47' },
        width: 2,
        style: 'dash',
      },
      left: {
        kind: 'line',
        color: { kind: 'srgb', value: '4472C4' },
        width: 3,
        style: 'solid',
      },
    });

    const partial = {
      top: {
        kind: 'line' as const,
        color: { kind: 'scheme' as const, value: 'accent1' as const },
        width: 2,
        style: 'dash' as const,
      },
      bottom: none,
    };
    tables[2]!.borders = partial;
    expect(tables[2]!.borders).toEqual(partial);
    expect(tables[2]!.rows[0]!.cells.map(({ borders }) => borders))
      .toEqual(Array(2).fill(partial));

    const reopened = await PptxDocument.open(await imported.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables[0]!.borders).toEqual(four(uniformLine));
    expect(reopenedTables[1]!.borders).toEqual(four(none));
    expect(reopenedTables[2]!.borders).toEqual(partial);
    expect(reopenedTables[2]!.rows[0]!.cells.map(({ borders }) => borders))
      .toEqual(Array(2).fill(partial));
  });

  it('projects and normalizes PptxGenJS table-level fill output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[{ text: 'Uniform A', options: {} }, { text: 'Uniform B', options: {} }]],
      {
        x: 0.5,
        y: 0.5,
        w: 4,
        h: 1,
        fill: { color: generated.SchemeColor.accent1, transparency: 25 },
      },
    );
    slide.addTable(
      [[{ text: 'Omitted A', options: {} }, { text: 'Omitted B', options: {} }]],
      { x: 0.5, y: 2, w: 4, h: 1 },
    );
    slide.addTable(
      [[
        { text: 'Inherited', options: {} },
        {
          text: 'Override',
          options: { fill: { color: 'FFFF00', transparency: 50 } },
        },
      ]],
      { x: 0.5, y: 3.5, w: 4, h: 1, fill: { color: '0000FF' } },
    );

    const imported = await importPptxGenJS(generated);
    const tables = imported.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(3);
    expect(tables[0]!.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 25,
    });
    expect(tables[0]!.rows[0]!.cells.map(({ fill }) => fill)).toEqual(Array(2).fill({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 25,
    }));
    expect(tables[1]!.fill).toBeUndefined();
    expect(tables[1]!.rows[0]!.cells.map(({ fill }) => fill)
      .every((fill) => fill === undefined)).toBe(true);
    expect(tables[2]!.fill).toBeUndefined();
    expect(tables[2]!.rows[0]!.cells.map(({ fill }) => fill)).toEqual([
      { kind: 'solid', color: { kind: 'srgb', value: '0000FF' } },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFFF00' },
        transparency: 50,
      },
    ]);

    tables[2]!.fill = { kind: 'none' };
    expect(tables[2]!.fill).toEqual({ kind: 'none' });
    expect(tables[2]!.rows[0]!.cells.map(({ fill }) => fill))
      .toEqual(Array(2).fill({ kind: 'none' }));

    const reopened = await PptxDocument.open(await imported.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables[0]!.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 25,
    });
    expect(reopenedTables[1]!.fill).toBeUndefined();
    expect(reopenedTables[2]!.fill).toEqual({ kind: 'none' });
    expect(reopenedTables[2]!.rows[0]!.cells.map(({ fill }) => fill))
      .toEqual(Array(2).fill({ kind: 'none' }));

    const xml = new TextDecoder().decode(
      reopened.opcPackage.requirePart(reopened.slides[0]!.partUri).bytes,
    );
    const physicalCells = [...xml.matchAll(/<a:tc\b[^>]*>([\s\S]*?)<\/a:tc>/g)]
      .map((match) => match[1]!);
    for (const cell of physicalCells.slice(-2)) {
      const properties = cell.match(/<a:tcPr\b[^>]*>([\s\S]*?)<\/a:tcPr>/);
      expect(properties).not.toBeNull();
      const withoutBorders = properties![1]!
        .replace(/<a:(ln[LRBT])\b[^>]*>[\s\S]*?<\/a:\1>/g, '')
        .replace(/<a:ln[LRBT]\b[^>]*\/>/g, '');
      const directFillChoices = [...withoutBorders.matchAll(
        /<a:(noFill|solidFill|gradFill|blipFill|pattFill|grpFill)\b/g,
      )].map((match) => match[1]);
      expect(directFillChoices).toEqual(['noFill']);
    }
  });

  it('projects and normalizes PptxGenJS table-level margin output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[{ text: 'Uniform A', options: {} }, { text: 'Uniform B', options: {} }]],
      { x: 0.5, y: 0.5, w: 4, h: 1, margin: [0.05, 0.1, 0.15, 0.2] },
    );
    slide.addTable(
      [[{ text: 'Omitted A', options: {} }, { text: 'Omitted B', options: {} }]],
      { x: 0.5, y: 2, w: 4, h: 1 },
    );
    slide.addTable(
      [[
        { text: 'Inherited', options: {} },
        { text: 'Override', options: { margin: 0 } },
      ]],
      { x: 0.5, y: 3.5, w: 4, h: 1, margin: 0.1 },
    );

    const imported = await importPptxGenJS(generated);
    const tables = imported.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(3);
    expect(tables[0]!.margins).toEqual({
      top: 3.6,
      right: 7.2,
      bottom: 10.8,
      left: 14.4,
    });
    expect(tables[0]!.rows[0]!.cells.map(({ margins }) => margins)).toEqual(Array(2).fill({
      top: 3.6,
      right: 7.2,
      bottom: 10.8,
      left: 14.4,
    }));
    expect(tables[1]!.margins).toEqual({
      top: 3.6,
      right: 7.2,
      bottom: 3.6,
      left: 7.2,
    });
    expect(tables[2]!.margins).toBeUndefined();
    expect(tables[2]!.rows[0]!.cells.map(({ margins }) => margins)).toEqual([
      { top: 7.2, right: 7.2, bottom: 7.2, left: 7.2 },
      { top: 0, right: 0, bottom: 0, left: 0 },
    ]);

    tables[2]!.margins = [1, 2, 3, 4];
    expect(tables[2]!.margins).toEqual({ top: 1, right: 2, bottom: 3, left: 4 });
    expect(tables[2]!.rows[0]!.cells.map(({ margins }) => margins)).toEqual(Array(2).fill({
      top: 1,
      right: 2,
      bottom: 3,
      left: 4,
    }));

    const reopened = await PptxDocument.open(await imported.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables[0]!.margins).toEqual({
      top: 3.6,
      right: 7.2,
      bottom: 10.8,
      left: 14.4,
    });
    expect(reopenedTables[1]!.margins).toEqual({
      top: 3.6,
      right: 7.2,
      bottom: 3.6,
      left: 7.2,
    });
    expect(reopenedTables[2]!.margins).toEqual({ top: 1, right: 2, bottom: 3, left: 4 });
    expect(reopenedTables[2]!.rows[0]!.cells.map(({ margins }) => margins)).toEqual(Array(2).fill({
      top: 1,
      right: 2,
      bottom: 3,
      left: 4,
    }));

    const xml = new TextDecoder().decode(
      reopened.opcPackage.requirePart(reopened.slides[0]!.partUri).bytes,
    );
    expect(xml.match(/<a:tcPr\b[^>]*marL="50800"[^>]*marR="25400"[^>]*marT="12700"[^>]*marB="38100"/g))
      .toHaveLength(2);
    expect(xml).not.toMatch(/<a:bodyPr\b[^>]*(?:marL|marR|marT|marB)=/);
  });

  it('projects and normalizes PptxGenJS table-level horizontal alignment output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    for (const [index, align] of ['left', 'center', 'right', 'justify'].entries()) {
      slide.addTable(
        [[{ text: `${align} A`, options: {} }, { text: `${align} B`, options: {} }]],
        {
          x: 0.5,
          y: 0.5 + (index * 0.8),
          w: 4,
          h: 0.6,
          align: align as 'left' | 'center' | 'right' | 'justify',
        },
      );
    }
    slide.addTable(
      [[{ text: 'Omitted A', options: {} }, { text: 'Omitted B', options: {} }]],
      { x: 5, y: 0.5, w: 4, h: 0.6 },
    );
    slide.addTable(
      [[
        { text: 'Inherited center', options: {} },
        { text: 'Right override', options: { align: 'right' } },
      ]],
      { x: 5, y: 1.5, w: 4, h: 0.6, align: 'center' },
    );

    const imported = await importPptxGenJS(generated);
    const tables = imported.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(6);
    expect(tables.slice(0, 4).map(({ horizontalAlignment }) => horizontalAlignment)).toEqual([
      'left',
      'center',
      'right',
      'justify',
    ]);
    expect(tables[4]!.horizontalAlignment).toBeUndefined();
    expect(tables[4]!.rows[0]!.cells.map(
      ({ horizontalAlignment }) => horizontalAlignment)).toEqual([undefined, undefined]);
    expect(tables[5]!.horizontalAlignment).toBeUndefined();
    expect(tables[5]!.rows[0]!.cells.map(
      ({ horizontalAlignment }) => horizontalAlignment)).toEqual(['center', 'right']);

    tables[5]!.horizontalAlignment = 'left';
    expect(tables[5]!.horizontalAlignment).toBe('left');
    expect(tables[5]!.rows[0]!.cells.map(
      ({ horizontalAlignment }) => horizontalAlignment)).toEqual(['left', 'left']);

    const reopened = await PptxDocument.open(await imported.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables.slice(0, 4).map(
      ({ horizontalAlignment }) => horizontalAlignment)).toEqual([
      'left',
      'center',
      'right',
      'justify',
    ]);
    expect(reopenedTables[4]!.horizontalAlignment).toBeUndefined();
    expect(reopenedTables[5]!.horizontalAlignment).toBe('left');
    expect(reopenedTables[5]!.rows[0]!.cells.map(
      ({ horizontalAlignment }) => horizontalAlignment)).toEqual(['left', 'left']);

    const xml = new TextDecoder().decode(
      reopened.opcPackage.requirePart(reopened.slides[0]!.partUri).bytes,
    );
    expect(xml.match(/<a:pPr\b[^>]* algn="l"/g)).toHaveLength(4);
    expect(xml.match(/<a:pPr\b[^>]* algn="ctr"/g)).toHaveLength(2);
    expect(xml.match(/<a:pPr\b[^>]* algn="r"/g)).toHaveLength(2);
    expect(xml.match(/<a:pPr\b[^>]* algn="just"/g)).toHaveLength(2);
    expect(xml).not.toMatch(/<a:tcPr\b[^>]* algn=/);
    expect(xml).not.toMatch(/<a:bodyPr\b[^>]* algn=/);
  });

  it('imports PptxGenJS table-cell text directions with exact four-value semantics', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[
        { text: 'Inherited', options: {} },
        { text: 'Horizontal', options: { textDirection: 'horz' } },
        { text: 'Vertical', options: { textDirection: 'vert' } },
        { text: 'Rotate 270', options: { textDirection: 'vert270' } },
        { text: 'Stacked', options: { textDirection: 'wordArtVert' } },
        { text: 'Invalid', options: { textDirection: 'eaVert' } },
      ]],
      { x: 0.5, y: 0.5, w: 12, h: 1, textDirection: 'vert270' },
    );
    slide.addTable(
      [[
        { text: 'Omitted', options: {} },
        { text: 'Explicit horizontal', options: { textDirection: 'horz' } },
      ]],
      { x: 0.5, y: 2, w: 12, h: 1 },
    );

    const document = await importPptxGenJS(generated);
    const tables = document.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(2);
    expect(tables[0]!.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'vert270',
      undefined,
      'vert',
      'vert270',
      'wordArtVert',
      undefined,
    ]);
    expect(tables[1]!.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      undefined,
      undefined,
    ]);
    expect(tables[0]!.rows[0]!.cells.map(({ text }) => text)).toEqual([
      'Inherited',
      'Horizontal',
      'Vertical',
      'Rotate 270',
      'Stacked',
      'Invalid',
    ]);
    expect(tables[1]!.rows[0]!.cells.map(({ text }) => text)).toEqual([
      'Omitted',
      'Explicit horizontal',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(xml.match(/<a:tcPr[^>]* vert="vert270"/g)).toHaveLength(2);
    expect(xml.match(/<a:tcPr[^>]* vert="vert"/g)).toHaveLength(1);
    expect(xml.match(/<a:tcPr[^>]* vert="wordArtVert"/g)).toHaveLength(1);
    expect(xml.match(/<a:tcPr[^>]* vert="eaVert"/g)).toHaveLength(1);
    expect(xml).not.toMatch(/<a:tcPr[^>]* vert="horz"/);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables[0]!.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'vert270',
      undefined,
      'vert',
      'vert270',
      'wordArtVert',
      undefined,
    ]);
    expect(reopenedTables[1]!.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      undefined,
      undefined,
    ]);
  });

  it('projects and normalizes PptxGenJS table-level text direction output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[{ text: 'Uniform A', options: {} }, { text: 'Uniform B', options: {} }]],
      { x: 0.5, y: 0.5, w: 4, h: 1, textDirection: 'vert270' },
    );
    slide.addTable(
      [[{ text: 'Horizontal A', options: {} }, { text: 'Horizontal B', options: {} }]],
      { x: 0.5, y: 2, w: 4, h: 1, textDirection: 'horz' },
    );
    slide.addTable(
      [[
        { text: 'Inherited stacked', options: {} },
        { text: 'Vertical override', options: { textDirection: 'vert' } },
      ]],
      { x: 0.5, y: 3.5, w: 4, h: 1, textDirection: 'wordArtVert' },
    );

    const imported = await importPptxGenJS(generated);
    const tables = imported.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(3);
    expect(tables[0]!.textDirection).toBe('vert270');
    expect(tables[1]!.textDirection).toBeUndefined();
    expect(tables[1]!.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      undefined,
      undefined,
    ]);
    expect(tables[2]!.textDirection).toBeUndefined();
    expect(tables[2]!.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'wordArtVert',
      'vert',
    ]);

    tables[2]!.textDirection = 'horz';
    expect(tables[2]!.textDirection).toBe('horz');
    expect(tables[2]!.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'horz',
      'horz',
    ]);
    tables[2]!.textDirection = 'wordArtVert';
    expect(tables[2]!.textDirection).toBe('wordArtVert');

    const reopened = await PptxDocument.open(await imported.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables[0]!.textDirection).toBe('vert270');
    expect(reopenedTables[1]!.textDirection).toBeUndefined();
    expect(reopenedTables[2]!.textDirection).toBe('wordArtVert');
    expect(reopenedTables[2]!.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'wordArtVert',
      'wordArtVert',
    ]);

    const xml = new TextDecoder().decode(
      reopened.opcPackage.requirePart(reopened.slides[0]!.partUri).bytes,
    );
    expect(xml.match(/<a:tcPr[^>]* vert="vert270"/g)).toHaveLength(2);
    expect(xml.match(/<a:tcPr[^>]* vert="wordArtVert"/g)).toHaveLength(2);
    expect(xml).not.toMatch(/<a:tcPr[^>]* vert="horz"/);
    expect(xml).not.toMatch(/<a:bodyPr[^>]* vert=/);
  });

  it('matches native table text direction creation to PptxGenJS final state', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable([[
      { text: 'Omitted', options: {} },
      { text: 'Horizontal', options: { textDirection: 'horz' } },
      { text: 'Vertical', options: { textDirection: 'vert' } },
      { text: 'Rotate 270', options: { textDirection: 'vert270' } },
      { text: 'Stacked', options: { textDirection: 'wordArtVert' } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 10,
      h: 1,
      colW: [2, 2, 2, 2, 2],
      rowH: 1,
      margin: 0.1,
      textDirection: 'vert270',
      valign: 'middle',
    });
    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes[0] as TableModel;

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeSlide = native.addSlide();
    const nativeTable = nativeSlide.addTable([[
      { text: 'Omitted' },
      { text: 'Horizontal', options: { textDirection: 'horz' } },
      { text: 'Vertical', options: { textDirection: 'vert' } },
      { text: 'Rotate 270', options: { textDirection: 'vert270' } },
      { text: 'Stacked', options: { textDirection: 'wordArtVert' } },
    ]], {
      x: inches(0.5),
      y: inches(0.5),
      width: inches(10),
      height: inches(1),
      columnWidths: inches(2),
      rowHeights: inches(1),
      margin: 7.2,
      textDirection: 'vert270',
      valign: 'middle',
    });
    const expectedTokens = ['vert270', undefined, 'vert', 'vert270', 'wordArtVert'];
    const expectedText = ['Omitted', 'Horizontal', 'Vertical', 'Rotate 270', 'Stacked'];
    const directDirectionTokens = (xml: string): (string | undefined)[] =>
      [...xml.matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g)]
        .map((match) => match[0]!
          .match(/<a:tcPr[^>]*\svert="([^"]+)"/)?.[1]);
    const slideXml = (document: PptxDocument): string => new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    const nativeXml = slideXml(native);
    const importedXml = slideXml(imported);

    expect(directDirectionTokens(nativeXml)).toEqual(expectedTokens);
    expect(directDirectionTokens(importedXml)).toEqual(expectedTokens);
    expect(nativeTable.rows[0]!.cells.map(({ textDirection }) => textDirection))
      .toEqual(expectedTokens);
    expect(importedTable.rows[0]!.cells.map(({ textDirection }) => textDirection))
      .toEqual(expectedTokens);
    expect(nativeTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(importedTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(nativeTable.transform).toMatchObject(importedTable.transform);
    expect(nativeTable.columnWidths).toEqual(importedTable.columnWidths);
    expect(nativeTable.rowHeights).toEqual(importedTable.rowHeights);
    expect(nativeTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(
      importedTable.rows[0]!.cells.map(({ margins }) => margins),
    );
    expect(nativeTable.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(
      importedTable.rows[0]!.cells.map(({ verticalAlignment }) => verticalAlignment),
    );
    expect(nativeTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(
      importedTable.rows[0]!.cells.map(({ borders }) => borders),
    );
    expect(nativeTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(
      importedTable.rows[0]!.cells.map(({ fill }) => fill),
    );
    for (const xml of [nativeXml, importedXml]) {
      expect(xml).not.toMatch(/<a:bodyPr[^>]*\svert=/);
    }

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect(directDirectionTokens(slideXml(reopenedNative))).toEqual(expectedTokens);
    expect(directDirectionTokens(slideXml(reopenedImported))).toEqual(expectedTokens);
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(nativeTable.rows);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(importedTable.rows);

    const generatedInvalid = new PptxGenJS();
    generatedInvalid.addSlide().addTable([[{
      text: 'PptxGenJS invalid east Asian vertical',
      options: { textDirection: 'eaVert' },
    }]], {
      x: 0.5,
      y: 0.5,
      w: 2,
      h: 1,
      colW: [2],
      rowH: 1,
    });
    const importedInvalid = await importPptxGenJS(generatedInvalid);
    expect(directDirectionTokens(slideXml(importedInvalid))).toEqual(['eaVert']);
    expect((importedInvalid.slides[0]!.shapes[0] as TableModel)
      .rows[0]!.cells[0]!.textDirection).toBeUndefined();

    const generatedInvalidTable = new PptxGenJS();
    generatedInvalidTable.addSlide().addTable([[{
      text: 'PptxGenJS invalid table direction',
      options: {},
    }]], {
      x: 0.5,
      y: 0.5,
      w: 2,
      h: 1,
      colW: [2],
      rowH: 1,
      textDirection: 'eaVert',
    });
    const importedInvalidTable = await importPptxGenJS(generatedInvalidTable);
    expect(directDirectionTokens(slideXml(importedInvalidTable))).toEqual(['eaVert']);
    expect((importedInvalidTable.slides[0]!.shapes[0] as TableModel)
      .rows[0]!.cells[0]!.textDirection).toBeUndefined();

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    const shapeCount = nativeSlide.shapes.length;
    expect(() => nativeSlide.addTable([[{
      text: 'Native invalid east Asian vertical',
      options: { textDirection: 'eaVert' as never },
    }]])).toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
    expect(nativeSlide.shapes).toHaveLength(shapeCount);
    expect(nativeSlide.shapes[0]).toBe(nativeTable);
    expect(() => nativeSlide.addTable([['Native invalid table direction']], {
      textDirection: 'eaVert' as never,
    })).toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
    expect(nativeSlide.shapes).toHaveLength(shapeCount);
    expect(nativeSlide.shapes[0]).toBe(nativeTable);
  });

  it('matches legal PptxGenJS table merges and native colspan and rowspan output', async () => {
    type MergeInput = string | {
      readonly text: string;
      readonly options?: {
        readonly colspan?: number;
        readonly rowspan?: number;
      };
    };
    const cases: readonly {
      readonly name: string;
      readonly rows: readonly (readonly MergeInput[])[];
      readonly region: {
        readonly rowIndex: number;
        readonly columnIndex: number;
        readonly rowspan: number;
        readonly colspan: number;
      };
    }[] = [
      {
        name: 'Horizontal',
        rows: [
          [{ text: 'Horizontal anchor', options: { colspan: 2 } }, 'Top right'],
          ['H 1', 'H 2', 'H 3'],
        ],
        region: { rowIndex: 0, columnIndex: 0, rowspan: 1, colspan: 2 },
      },
      {
        name: 'Vertical',
        rows: [
          [{ text: 'Vertical anchor', options: { rowspan: 2 } }, 'Top right'],
          ['Bottom right'],
        ],
        region: { rowIndex: 0, columnIndex: 0, rowspan: 2, colspan: 1 },
      },
      {
        name: 'Rectangular',
        rows: [
          [{ text: 'Rectangle anchor', options: { colspan: 2, rowspan: 2 } }, 'Top right'],
          ['Bottom right'],
          ['R 1', 'R 2', 'R 3'],
        ],
        region: { rowIndex: 0, columnIndex: 0, rowspan: 2, colspan: 2 },
      },
      {
        name: 'Offset',
        rows: [
          ['Top left', { text: 'Offset anchor', options: { colspan: 2, rowspan: 2 } }, 'Top right'],
          ['Bottom left', 'Bottom right'],
          ['O 1', 'O 2', 'O 3', 'O 4'],
        ],
        region: { rowIndex: 0, columnIndex: 1, rowspan: 2, colspan: 2 },
      },
      {
        name: 'Full row',
        rows: [
          [{ text: 'Full row anchor', options: { colspan: 3 } }],
          ['F 1', 'F 2', 'F 3'],
        ],
        region: { rowIndex: 0, columnIndex: 0, rowspan: 1, colspan: 3 },
      },
    ];
    const generated = new PptxGenJS();
    generated.layout = 'LAYOUT_WIDE';
    const native = PptxDocument.create({ slideSize: 'wide' });
    for (const fixture of cases) {
      generated.addSlide().addTable(structuredClone(fixture.rows), {
        x: 0.5,
        y: 0.5,
        w: 6,
        h: 2,
        margin: 0,
      });
      native.addSlide().addTable(structuredClone(fixture.rows), {
        x: inches(0.5),
        y: inches(0.5),
        width: inches(6),
        height: inches(2),
        margin: 0,
      });
    }
    expect(generated.version).toBe('4.0.1');
    const imported = await importPptxGenJS(generated);
    const mergeTokenMatrices = (document: PptxDocument, slideIndex: number) => {
      const table = slideXml(document, slideIndex)
        .match(/<a:tbl(?:\s[^>]*)?>[\s\S]*?<\/a:tbl>/)?.[0];
      if (!table) throw new Error(`Table ${slideIndex} was not found`);
      return [...table.matchAll(/<a:tr(?:\s[^>]*)?>([\s\S]*?)<\/a:tr>/g)]
        .map((row) => [...row[1]!.matchAll(/<a:tc(?:\s[^>]*)?>/g)].map((cell) => {
          const attribute = (name: string): string | undefined =>
            cell[0].match(new RegExp(`\\s${name}="([^"]+)"`))?.[1];
          const boolean = (name: string): boolean => {
            const value = attribute(name);
            return value === '1' || value === 'true';
          };
          return {
            rowSpan: Number(attribute('rowSpan') ?? 1),
            gridSpan: Number(attribute('gridSpan') ?? 1),
            vMerge: boolean('vMerge'),
            hMerge: boolean('hMerge'),
          };
        }));
    };

    for (const [index, fixture] of cases.entries()) {
      const importedTable = imported.slides[index]!.shapes[0] as TableModel;
      const nativeTable = native.slides[index]!.shapes[0] as TableModel;
      expect(importedTable).toBeInstanceOf(TableModel);
      expect(nativeTable).toBeInstanceOf(TableModel);
      expect(importedTable.rows.map(({ cells }) => cells.length))
        .toEqual(nativeTable.rows.map(({ cells }) => cells.length));
      expect(importedTable.rows.map(({ cells }) => cells.map(({ text }) => text)))
        .toEqual(nativeTable.rows.map(({ cells }) => cells.map(({ text }) => text)));
      expect(importedTable.mergeRegions).toEqual([fixture.region]);
      expect(nativeTable.mergeRegions).toEqual([fixture.region]);
      expect(mergeTokenMatrices(imported, index)).toEqual(mergeTokenMatrices(native, index));

      importedTable.unmergeCell(
        fixture.region.rowIndex + fixture.region.rowspan - 1,
        fixture.region.columnIndex + fixture.region.colspan - 1,
      );
      expect(importedTable.mergeRegions).toEqual([]);
      importedTable.mergeCells(
        fixture.region.rowIndex,
        fixture.region.columnIndex,
        fixture.region.rowspan,
        fixture.region.colspan,
      );
      expect(importedTable.mergeRegions).toEqual([fixture.region]);
      expect(mergeTokenMatrices(imported, index)).toEqual(mergeTokenMatrices(native, index));
    }

    const reopenedImported = await PptxDocument.open(await imported.write());
    const reopenedNative = await PptxDocument.open(await native.write());
    for (const [index, fixture] of cases.entries()) {
      const importedTable = reopenedImported.slides[index]!.shapes[0] as TableModel;
      const nativeTable = reopenedNative.slides[index]!.shapes[0] as TableModel;
      expect(importedTable.mergeRegions).toEqual([fixture.region]);
      expect(nativeTable.mergeRegions).toEqual([fixture.region]);
      expect(importedTable.rows.map(({ cells }) => cells.map(({ text }) => text)))
        .toEqual(nativeTable.rows.map(({ cells }) => cells.map(({ text }) => text)));
      expect(mergeTokenMatrices(reopenedImported, index))
        .toEqual(mergeTokenMatrices(reopenedNative, index));
    }
  });

  it('edits rows and columns in a public PptxGenJS table without losing survivor state', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const survivorLink: { url: string; tooltip: string; _rId?: number } = {
      url: 'https://table-structure.example?a=1&b=2',
      tooltip: 'PptxGenJS survivor',
    };
    generated.addSlide().addTable([
      [
        {
          text: 'Merge anchor',
          options: { colspan: 2, rowspan: 2, fill: { color: 'DDEEFF' } },
        },
        { text: 'R0C2', options: {} },
        { text: 'R0C3', options: {} },
      ],
      [
        { text: 'R1C2', options: {} },
        { text: 'R1C3', options: {} },
      ],
      [
        { text: 'R2C0', options: {} },
        { text: 'R2C1', options: {} },
        {
          text: [
            { text: 'Styled ', options: { bold: true, color: '1F4E78' } },
            { text: 'linked survivor', options: { italic: true } },
          ],
          options: {
            fill: { color: 'F4B183' },
            hyperlink: survivorLink,
          },
        },
        { text: 'R2C3', options: {} },
      ],
      [
        { text: 'R3C0', options: {} },
        { text: 'R3C1', options: {} },
        { text: 'R3C2', options: {} },
        { text: 'R3C3', options: {} },
      ],
    ], {
      x: 1,
      y: 1,
      w: 10,
      h: 5,
      colW: [1, 2, 3, 4],
      rowH: [0.5, 1, 1.5, 2],
      margin: 0.1,
    });

    const imported = await importPptxGenJS(generated);
    const table = imported.slides[0]!.shapes[0] as TableModel;
    expect(table).toBeInstanceOf(TableModel);
    expect(table.mergeRegions).toEqual([
      { rowIndex: 0, columnIndex: 0, rowspan: 2, colspan: 2 },
    ]);
    expect(table.columnWidths).toEqual([inches(1), inches(2), inches(3), inches(4)]);
    expect(table.rowHeights).toEqual([
      inches(0.5),
      inches(1),
      inches(1.5),
      inches(2),
    ]);
    const relationshipId = imported.slides[0]!.relationships.find(
      ({ type }) => type.endsWith('/hyperlink'),
    )!.id;

    table.insertRows(1, { rowHeights: inches(0.25) });
    table.insertColumns(1, { columnWidths: inches(0.5) });
    table.setCellText(1, 1, 'Inserted hidden cell');
    table.deleteRows(4);
    table.deleteColumns(4);

    const assertEditedState = (candidate: TableModel): void => {
      expect(candidate.mergeRegions).toEqual([
        { rowIndex: 0, columnIndex: 0, rowspan: 3, colspan: 3 },
      ]);
      expect(candidate.columnWidths).toEqual([
        inches(1),
        inches(0.5),
        inches(2),
        inches(3),
      ]);
      expect(candidate.rowHeights).toEqual([
        inches(0.5),
        inches(0.25),
        inches(1),
        inches(1.5),
      ]);
      expect(candidate.transform.width).toBe(inches(6.5));
      expect(candidate.transform.height).toBe(inches(3.25));
      expect(candidate.rows[1]!.cells[1]!.text).toBe('Inserted hidden cell');
      const survivor = candidate.rows[3]!.cells[3]!;
      expect(survivor.text).toBe('Styled linked survivor');
      expect(survivor.fill).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: 'F4B183' },
      });
      expect(survivor.richText[0]!.runs.map(({ style }) => style)).toEqual([
        expect.objectContaining({
          bold: true,
          color: { kind: 'srgb', value: '1F4E78' },
          hyperlink: {
            url: 'https://table-structure.example?a=1&b=2',
            tooltip: 'PptxGenJS survivor',
          },
        }),
        expect.objectContaining({
          italic: true,
          hyperlink: {
            url: 'https://table-structure.example?a=1&b=2',
            tooltip: 'PptxGenJS survivor',
          },
        }),
      ]);
    };
    assertEditedState(table);
    expect(imported.slides[0]!.relationships.find(({ id }) => id === relationshipId))
      .toMatchObject({
        target: 'https://table-structure.example?a=1&b=2',
        targetMode: 'External',
      });

    const reopened = await PptxDocument.open(await imported.write());
    assertEditedState(reopened.slides[0]!.shapes[0] as TableModel);
    expect(reopened.slides[0]!.relationships.find(({ id }) => id === relationshipId))
      .toMatchObject({
        target: 'https://table-structure.example?a=1&b=2',
        targetMode: 'External',
      });
  }, 20_000);

  it('locks PptxGenJS colspan and rowspan mutation and invalid-output differences', async () => {
    const callerCell = {
      text: 'Caller anchor',
      options: { colspan: 2, rowspan: 2 } as Record<string, unknown>,
    };
    const callerRows = [[callerCell, 'Top right'], ['Bottom right']];
    const callerBefore = structuredClone(callerRows);
    const mutating = new PptxGenJS();
    mutating.addSlide().addTable(callerRows, { x: 1, y: 1, w: 4, h: 2 });
    expect(callerCell.options).toHaveProperty('border');
    await mutating.write({ outputType: 'nodebuffer', compression: true });
    expect(callerRows).not.toEqual(callerBefore);
    expect(callerCell.options).toMatchObject({
      colspan: 2,
      rowspan: 2,
      color: '000000',
      fontSize: 12,
      _lineIdx: 0,
    });

    type InvalidInput = string | {
      readonly text: string;
      readonly options?: {
        readonly colspan?: number;
        readonly rowspan?: number;
      };
    };
    const invalidCases: readonly {
      readonly name: string;
      readonly rows: readonly (readonly InvalidInput[])[];
      readonly physicalRows: readonly number[];
      readonly gridColumns: number;
      readonly emittedToken?: string;
    }[] = [
      {
        name: 'lopsided row',
        rows: [['A', 'B'], ['C']],
        physicalRows: [2, 1],
        gridColumns: 2,
      },
      {
        name: 'negative colspan',
        rows: [[{ text: 'Negative', options: { colspan: -1 } }, 'B']],
        physicalRows: [2],
        gridColumns: 0,
      },
      {
        name: 'fractional rowspan',
        rows: [
          [{ text: 'Fractional', options: { rowspan: 1.5 } }, 'B'],
          ['C', 'D'],
        ],
        physicalRows: [2, 3],
        gridColumns: 2,
        emittedToken: 'rowSpan="1.5"',
      },
      {
        name: 'out-of-bounds rowspan',
        rows: [
          [{ text: 'Out of bounds', options: { rowspan: 3 } }, 'B'],
          ['C'],
        ],
        physicalRows: [2, 2],
        gridColumns: 2,
        emittedToken: 'rowSpan="3"',
      },
    ];

    for (const fixture of invalidCases) {
      const generated = new PptxGenJS();
      generated.addSlide().addTable(structuredClone(fixture.rows), {
        x: 1,
        y: 1,
        w: 4,
        h: 2,
      });
      const imported = await importPptxGenJS(generated);
      const table = imported.slides[0]!.shapes[0] as TableModel;
      const xml = slideXml(imported, 0);
      expect(table.rows.map(({ cells }) => cells.length), fixture.name)
        .toEqual(fixture.physicalRows);
      expect(table.mergeRegions, fixture.name).toBeUndefined();
      expect(xml.match(/<a:gridCol\b/g)?.length ?? 0, fixture.name)
        .toBe(fixture.gridColumns);
      if (fixture.emittedToken !== undefined) {
        expect(xml, fixture.name).toContain(fixture.emittedToken);
      }

      const native = PptxDocument.create();
      const nativeSlide = native.addSlide();
      const nativeRows = structuredClone(fixture.rows);
      const nativeRowsBefore = structuredClone(nativeRows);
      const partBefore = native.opcPackage.requirePart(nativeSlide.partUri).bytes;
      const journalBefore = [...native.opcPackage.mutations];
      expect(() => nativeSlide.addTable(nativeRows), fixture.name).toThrow();
      expect(nativeRows, fixture.name).toEqual(nativeRowsBefore);
      expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes, fixture.name)
        .toEqual(partBefore);
      expect(native.opcPackage.mutations, fixture.name).toEqual(journalBefore);
      expect(nativeSlide.shapes, fixture.name).toEqual([]);
    }
  });

  it('imports PptxGenJS table fit-like runtime options as fit-less cells', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[
        { text: 'Omitted', options: {} },
        { text: 'Fit none', options: { fit: 'none' } },
        { text: 'Fit shrink', options: { fit: 'shrink' } },
        { text: 'Fit resize', options: { fit: 'resize' } },
        { text: 'Auto fit', options: { autoFit: true } },
        { text: 'Shrink text', options: { shrinkText: true } },
        {
          text: 'Conflicting',
          options: { fit: 'resize', autoFit: true, shrinkText: true, textDirection: 'vert' },
        },
      ]],
      {
        x: 0.5,
        y: 0.5,
        w: 12,
        h: 1,
        fit: 'resize',
        autoFit: true,
        shrinkText: true,
        textDirection: 'vert270',
      },
    );

    const document = await importPptxGenJS(generated);
    const table = document.slides[0]!.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(table?.rows[0]!.cells.map(({ textFit }) => textFit)).toEqual(
      Array(7).fill(undefined),
    );
    expect(table?.rows[0]!.cells.map(({ text }) => text)).toEqual([
      'Omitted',
      'Fit none',
      'Fit shrink',
      'Fit resize',
      'Auto fit',
      'Shrink text',
      'Conflicting',
    ]);
    expect(table?.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(xml).not.toMatch(/<a:(?:noAutofit|normAutofit|spAutoFit)\b/);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTable?.rows[0]!.cells.map(({ textFit }) => textFit)).toEqual(
      Array(7).fill(undefined),
    );
    expect(reopenedTable?.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert270',
      'vert',
    ]);

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeSlide = native.addSlide();
    const nativeTable = nativeSlide.addTable([[
      { text: 'Omitted' },
      { text: 'None', options: { fit: 'none' } },
      { text: 'Shrink', options: { fit: 'shrink' } },
      { text: 'Resize', options: { fit: 'resize', textDirection: 'vert' } },
    ]], {
      columnWidths: inches(2),
      rowHeights: inches(1),
    });
    expect(nativeTable.rows[0]!.cells.map(({ textFit }) => textFit)).toEqual([
      undefined,
      undefined,
      'shrink',
      'resize',
    ]);
    expect(nativeTable.rows[0]!.cells[3]!.textDirection).toBe('vert');

    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(nativeSlide.partUri).bytes,
    );
    expect(nativeXml.match(/<a:normAutofit\/>/g)).toHaveLength(1);
    expect(nativeXml.match(/<a:spAutoFit\/>/g)).toHaveLength(1);
    expect(nativeXml).not.toContain('<a:noAutofit/>');

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes;
    const invalidJournal = [...native.opcPackage.mutations];
    expect(() => nativeSlide.addTable([[
      { text: 'Invalid', options: { fit: 'SHRINK' as never } },
    ]])).toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
    expect(nativeSlide.shapes[0]).toBe(nativeTable);
  });

  it('imports PptxGenJS table-cell vertical alignments from direct cell anchors', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[
        { text: 'Inherited bottom', options: {} },
        { text: 'Top', options: { valign: 'top' } },
        { text: 'Middle', options: { valign: 'middle' } },
        { text: 'Bottom', options: { valign: 'bottom' } },
        { text: 'Invalid mid', options: { valign: 'mid' } },
        { text: 'Invalid distributed', options: { valign: 'distributed' } },
      ]],
      { x: 0.5, y: 0.5, w: 12, h: 1, valign: 'bottom' },
    );
    slide.addTable(
      [[{ text: 'Inherited top', options: {} }]],
      { x: 0.5, y: 2, w: 3, h: 1, valign: 'top' },
    );
    slide.addTable(
      [[{ text: 'Inherited middle', options: {} }]],
      { x: 4, y: 2, w: 3, h: 1, valign: 'middle' },
    );
    slide.addTable(
      [[{ text: 'Omitted direct alignment', options: {} }]],
      { x: 7.5, y: 2, w: 3, h: 1 },
    );

    const document = await importPptxGenJS(generated);
    const tables = document.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(4);
    expect(tables[0]!.rows[0]!.cells.map(({ verticalAlignment }) => verticalAlignment)).toEqual([
      'bottom',
      'top',
      'middle',
      'bottom',
      undefined,
      undefined,
    ]);
    expect(tables.slice(1).map((table) => table.rows[0]!.cells[0]!.verticalAlignment)).toEqual([
      'top',
      'middle',
      undefined,
    ]);
    expect(tables.flatMap(({ rows }) => rows[0]!.cells.map(({ text }) => text))).toEqual([
      'Inherited bottom',
      'Top',
      'Middle',
      'Bottom',
      'Invalid mid',
      'Invalid distributed',
      'Inherited top',
      'Inherited middle',
      'Omitted direct alignment',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(xml.match(/<a:tcPr[^>]* anchor="t"/g)).toHaveLength(2);
    expect(xml.match(/<a:tcPr[^>]* anchor="ctr"/g)).toHaveLength(2);
    expect(xml.match(/<a:tcPr[^>]* anchor="b"/g)).toHaveLength(2);
    expect(xml.match(/<a:tcPr[^>]* anchor="mid"/g)).toHaveLength(1);
    expect(xml.match(/<a:tcPr[^>]* anchor="distributed"/g)).toHaveLength(1);
    expect(xml).not.toMatch(/<a:bodyPr[^>]* anchor=/);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables[0]!.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual([
      'bottom',
      'top',
      'middle',
      'bottom',
      undefined,
      undefined,
    ]);
    expect(reopenedTables.slice(1).map(
      (table) => table.rows[0]!.cells[0]!.verticalAlignment)).toEqual([
      'top',
      'middle',
      undefined,
    ]);
  });

  it('matches native table-level valign creation to supported PptxGenJS output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable([[
      { text: 'Inherited string', options: {} },
      { text: 'Inherited object', options: {} },
      { text: 'Top', options: { valign: 'top' } },
      { text: 'Middle', options: { valign: 'middle' } },
      { text: 'Bottom', options: { valign: 'bottom' } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 8,
      h: 1,
      colW: [1.6, 1.6, 1.6, 1.6, 1.6],
      rowH: 1,
      valign: 'middle',
    });
    generated.addSlide().addTable([[
      { text: 'Uniform one', options: {} },
      { text: 'Uniform two', options: {} },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 4,
      h: 1,
      colW: [2, 2],
      rowH: 1,
      valign: 'middle',
    });
    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    const uniformImportedTable = imported.slides[1]!.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(importedTable).toBeInstanceOf(TableModel);
    expect(uniformImportedTable).toBeInstanceOf(TableModel);

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeSlide = native.addSlide();
    const nativeTable = nativeSlide.addTable([[
      'Inherited string',
      { text: 'Inherited object' },
      { text: 'Top', options: { valign: 'top' } },
      { text: 'Middle', options: { valign: 'middle' } },
      { text: 'Bottom', options: { valign: 'bottom' } },
    ]], {
      x: inches(0.5),
      y: inches(0.5),
      columnWidths: inches(1.6),
      rowHeights: inches(1),
      valign: 'middle',
    });
    const expectedAlignments = ['middle', 'middle', 'top', 'middle', 'bottom'];
    const expectedText = ['Inherited string', 'Inherited object', 'Top', 'Middle', 'Bottom'];
    expect(nativeTable.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(expectedAlignments);
    expect(importedTable!.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(expectedAlignments);
    expect(importedTable!.verticalAlignment).toBeUndefined();
    expect(uniformImportedTable!.verticalAlignment).toBe('middle');
    expect(nativeTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(importedTable!.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(nativeTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(
      importedTable!.rows[0]!.cells.map(({ margins }) => margins),
    );
    expect(nativeTable.transform).toMatchObject(importedTable!.transform);
    expect(nativeTable.columnWidths).toEqual(importedTable!.columnWidths);
    expect(nativeTable.rowHeights).toEqual(importedTable!.rowHeights);

    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(nativeSlide.partUri).bytes,
    );
    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    for (const xml of [nativeXml, importedXml]) {
      const properties = [...xml.matchAll(/<a:tcPr([^>]*)>/g)]
        .map((match) => match[1]!);
      expect(properties).toHaveLength(5);
      expect(properties.map((attributes) =>
        attributes.match(/\sanchor="([^"]+)"/)?.[1])).toEqual([
        'ctr',
        'ctr',
        't',
        'ctr',
        'b',
      ]);
      expect(xml).not.toMatch(/<a:bodyPr[^>]*\sanchor=/);
    }

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(expectedAlignments);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(expectedAlignments);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).verticalAlignment)
      .toBeUndefined();
    expect((reopenedImported.slides[1]!.shapes[0] as TableModel).verticalAlignment)
      .toBe('middle');

    importedTable!.verticalAlignment = 'bottom';
    expect(importedTable!.verticalAlignment).toBe('bottom');
    expect(importedTable!.rows[0]!.cells.every(
      ({ verticalAlignment }) => verticalAlignment === 'bottom',
    )).toBe(true);
    const normalizedImported = await PptxDocument.open(await imported.write());
    const normalizedImportedTable = normalizedImported.slides[0]!.shapes[0] as TableModel;
    expect(normalizedImportedTable.verticalAlignment).toBe('bottom');
    expect(normalizedImportedTable.rows[0]!.cells.every(
      ({ verticalAlignment }) => verticalAlignment === 'bottom',
    )).toBe(true);

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    expect(() => nativeSlide.addTable([[{
      text: 'Invalid mid',
      options: { valign: 'mid' as never },
    }]])).toThrow(TypeError);
    expect(() => nativeSlide.addTable(
      [['Invalid distributed']],
      { valign: 'distributed' as never },
    )).toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
  });

  it('matches native table-cell horizontal alignment to PptxGenJS final state', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable([[
      { text: 'Default', options: {} },
      { text: 'Left', options: { align: 'left' } },
      { text: 'Center', options: { align: 'center' } },
      { text: 'Right', options: { align: 'right' } },
      { text: 'Justify this sentence', options: { align: 'justify' } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 10,
      h: 1,
      colW: [2, 2, 2, 2, 2],
      rowH: 1,
      margin: 0.1,
      valign: 'middle',
    });
    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes[0] as TableModel;

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeSlide = native.addSlide();
    const nativeTable = nativeSlide.addTable([[
      { text: 'Default' },
      { text: 'Left', options: { align: 'left' } },
      { text: 'Center', options: { align: 'center' } },
      { text: 'Right', options: { align: 'right' } },
      { text: 'Justify this sentence', options: { align: 'justify' } },
    ]], {
      x: inches(0.5),
      y: inches(0.5),
      width: inches(10),
      height: inches(1),
      columnWidths: inches(2),
      rowHeights: inches(1),
      margin: 7.2,
      valign: 'middle',
    });
    const expectedTokens = [undefined, 'l', 'ctr', 'r', 'just'];
    const expectedAlignments = [undefined, 'left', 'center', 'right', 'justify'];
    const expectedText = ['Default', 'Left', 'Center', 'Right', 'Justify this sentence'];
    const directAlignmentTokens = (xml: string): (string | undefined)[] =>
      [...xml.matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g)]
        .map((match) => match[0]!
          .match(/<a:pPr[^>]*\salgn="([^"]+)"/)?.[1]);
    const slideXml = (document: PptxDocument): string => new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    const nativeXml = slideXml(native);
    const importedXml = slideXml(imported);

    expect(directAlignmentTokens(nativeXml)).toEqual(expectedTokens);
    expect(directAlignmentTokens(importedXml)).toEqual(expectedTokens);
    expect(nativeTable.rows[0]!.cells.map(({ horizontalAlignment }) =>
      horizontalAlignment)).toEqual(expectedAlignments);
    expect(importedTable.rows[0]!.cells.map(({ horizontalAlignment }) =>
      horizontalAlignment)).toEqual(expectedAlignments);
    expect(nativeTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(importedTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(nativeTable.transform).toMatchObject(importedTable.transform);
    expect(nativeTable.columnWidths).toEqual(importedTable.columnWidths);
    expect(nativeTable.rowHeights).toEqual(importedTable.rowHeights);
    expect(nativeTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(
      importedTable.rows[0]!.cells.map(({ margins }) => margins),
    );
    expect(nativeTable.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(
      importedTable.rows[0]!.cells.map(({ verticalAlignment }) => verticalAlignment),
    );
    expect(nativeTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(
      importedTable.rows[0]!.cells.map(({ borders }) => borders),
    );
    expect(nativeTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(
      importedTable.rows[0]!.cells.map(({ fill }) => fill),
    );
    for (const xml of [nativeXml, importedXml]) {
      expect(xml).not.toMatch(/<a:tcPr[^>]*\salgn=/);
      expect(xml).not.toMatch(/<a:bodyPr[^>]*\salgn=/);
    }

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect(directAlignmentTokens(slideXml(reopenedNative))).toEqual(expectedTokens);
    expect(directAlignmentTokens(slideXml(reopenedImported))).toEqual(expectedTokens);
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows[0]!.cells.map(
      ({ horizontalAlignment }) => horizontalAlignment)).toEqual(expectedAlignments);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows[0]!.cells.map(
      ({ horizontalAlignment }) => horizontalAlignment)).toEqual(expectedAlignments);
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(nativeTable.rows);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(importedTable.rows);

    const generatedInvalid = new PptxGenJS();
    generatedInvalid.addSlide().addTable([[{
      text: 'PptxGenJS invalid distributed',
      options: { align: 'dist' },
    }]], {
      x: 0.5,
      y: 0.5,
      w: 2,
      h: 1,
      colW: [2],
      rowH: 1,
    });
    const importedInvalid = await importPptxGenJS(generatedInvalid);
    expect(directAlignmentTokens(slideXml(importedInvalid))).toEqual([undefined]);
    expect((importedInvalid.slides[0]!.shapes[0] as TableModel)
      .rows[0]!.cells[0]!.horizontalAlignment).toBeUndefined();

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    const shapeCount = nativeSlide.shapes.length;
    expect(() => nativeSlide.addTable([[{
      text: 'Native invalid distributed',
      options: { align: 'dist' as never },
    }]])).toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
    expect(nativeSlide.shapes).toHaveLength(shapeCount);
    expect(nativeSlide.shapes[0]).toBe(nativeTable);
  });

  it('matches native table horizontal alignment to PptxGenJS final state', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable([[
      { text: 'Inherited string', options: {} },
      { text: 'Inherited object', options: {} },
      { text: 'Inherited undefined', options: { align: undefined } },
      { text: 'Left override', options: { align: 'left' } },
      { text: 'Right override', options: { align: 'right' } },
      { text: 'Justify override', options: { align: 'justify' } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 9,
      h: 1,
      colW: [1.5, 1.5, 1.5, 1.5, 1.5, 1.5],
      rowH: 1,
      align: 'center',
      margin: 0.1,
      valign: 'middle',
    });
    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes[0] as TableModel;

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeSlide = native.addSlide();
    const nativeTable = nativeSlide.addTable([[
      'Inherited string',
      { text: 'Inherited object' },
      { text: 'Inherited undefined', options: { align: undefined } as never },
      { text: 'Left override', options: { align: 'left' } },
      { text: 'Right override', options: { align: 'right' } },
      { text: 'Justify override', options: { align: 'justify' } },
    ]], {
      x: inches(0.5),
      y: inches(0.5),
      columnWidths: inches(1.5),
      rowHeights: inches(1),
      align: 'center',
      margin: 7.2,
      valign: 'middle',
    });
    const expectedTokens = ['ctr', 'ctr', 'ctr', 'l', 'r', 'just'];
    const expectedAlignments = ['center', 'center', 'center', 'left', 'right', 'justify'];
    const expectedText = [
      'Inherited string',
      'Inherited object',
      'Inherited undefined',
      'Left override',
      'Right override',
      'Justify override',
    ];
    const directAlignmentTokens = (xml: string): (string | undefined)[] =>
      [...xml.matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g)]
        .map((match) => match[0]!
          .match(/<a:pPr[^>]*\salgn="([^"]+)"/)?.[1]);
    const slideXml = (document: PptxDocument): string => new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    const nativeXml = slideXml(native);
    const importedXml = slideXml(imported);

    expect(directAlignmentTokens(nativeXml)).toEqual(expectedTokens);
    expect(directAlignmentTokens(importedXml)).toEqual(expectedTokens);
    expect(nativeTable.rows[0]!.cells.map(({ horizontalAlignment }) =>
      horizontalAlignment)).toEqual(expectedAlignments);
    expect(importedTable.rows[0]!.cells.map(({ horizontalAlignment }) =>
      horizontalAlignment)).toEqual(expectedAlignments);
    expect(nativeTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(importedTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(nativeTable.transform).toMatchObject(importedTable.transform);
    expect(nativeTable.columnWidths).toEqual(importedTable.columnWidths);
    expect(nativeTable.rowHeights).toEqual(importedTable.rowHeights);
    expect(nativeTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(
      importedTable.rows[0]!.cells.map(({ margins }) => margins),
    );
    expect(nativeTable.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(
      importedTable.rows[0]!.cells.map(({ verticalAlignment }) => verticalAlignment),
    );
    expect(nativeTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(
      importedTable.rows[0]!.cells.map(({ borders }) => borders),
    );
    expect(nativeTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(
      importedTable.rows[0]!.cells.map(({ fill }) => fill),
    );
    for (const xml of [nativeXml, importedXml]) {
      expect(xml).not.toMatch(/<a:tcPr[^>]*\salgn=/);
      expect(xml).not.toMatch(/<a:bodyPr[^>]*\salgn=/);
    }

    const importedSlidePartUri = imported.slides[0]!.partUri;
    const nonSlideParts = new Map(imported.opcPackage.parts
      .filter(({ uri }) => uri !== importedSlidePartUri)
      .map(({ uri, bytes }) => [uri, bytes.slice()]));
    const importedText = importedTable.rows[0]!.cells.map(({ text }) => text);
    const importedMargins = importedTable.rows[0]!.cells.map(({ margins }) => margins);
    const importedVerticalAlignments = importedTable.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment);
    const importedBorders = importedTable.rows[0]!.cells.map(({ borders }) => borders);
    const importedFills = importedTable.rows[0]!.cells.map(({ fill }) => fill);
    const importedTransform = importedTable.transform;
    const importedColumnWidths = importedTable.columnWidths;
    const importedRowHeights = importedTable.rowHeights;

    importedTable.setCellHorizontalAlignment(0, 0, 'right');
    importedTable.setCellHorizontalAlignment(0, 3, undefined);
    importedTable.setCellHorizontalAlignment(0, 5, 'center');
    const editedTokens = ['r', 'ctr', 'ctr', undefined, 'r', 'ctr'];
    const editedAlignments = ['right', 'center', 'center', undefined, 'right', 'center'];
    expect(directAlignmentTokens(slideXml(imported))).toEqual(editedTokens);
    expect(importedTable.rows[0]!.cells.map(({ horizontalAlignment }) =>
      horizontalAlignment)).toEqual(editedAlignments);
    expect(importedTable.rows[0]!.cells.map(({ text }) => text)).toEqual(importedText);
    expect(importedTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(importedMargins);
    expect(importedTable.rows[0]!.cells.map(({ verticalAlignment }) =>
      verticalAlignment)).toEqual(importedVerticalAlignments);
    expect(importedTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(importedBorders);
    expect(importedTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(importedFills);
    expect(importedTable.transform).toEqual(importedTransform);
    expect(importedTable.columnWidths).toEqual(importedColumnWidths);
    expect(importedTable.rowHeights).toEqual(importedRowHeights);
    for (const [uri, bytes] of nonSlideParts) {
      expect(imported.opcPackage.requirePart(uri).bytes).toEqual(bytes);
    }

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect(directAlignmentTokens(slideXml(reopenedNative))).toEqual(expectedTokens);
    expect(directAlignmentTokens(slideXml(reopenedImported))).toEqual(editedTokens);
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(nativeTable.rows);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(importedTable.rows);
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows[0]!.cells.map(
      ({ horizontalAlignment }) => horizontalAlignment)).toEqual(expectedAlignments);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows[0]!.cells.map(
      ({ horizontalAlignment }) => horizontalAlignment)).toEqual(editedAlignments);

    const generatedInvalid = new PptxGenJS();
    generatedInvalid.addSlide().addTable([[
      { text: 'Invalid inherited', options: {} },
      { text: 'Right override', options: { align: 'right' } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 4,
      h: 1,
      colW: [2, 2],
      rowH: 1,
      align: 'dist',
    });
    const importedInvalid = await importPptxGenJS(generatedInvalid);
    expect(directAlignmentTokens(slideXml(importedInvalid))).toEqual([
      undefined,
      'r',
    ]);

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    const shapeCount = nativeSlide.shapes.length;
    expect(() => nativeSlide.addTable([
      ['Invalid inherited', { text: 'Right override', options: { align: 'right' } }],
    ], {
      align: 'dist' as never,
    })).toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
    expect(nativeSlide.shapes).toHaveLength(shapeCount);
    expect(nativeSlide.shapes[0]).toBe(nativeTable);
  });

  it('imports PptxGenJS table-cell margins from direct cell properties', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[{ text: 'Omitted defaults', options: {} }]],
      { x: 0.2, y: 0.2, w: 2, h: 0.5 },
    );
    slide.addTable(
      [[{ text: 'Table zero', options: {} }]],
      { x: 0.2, y: 1, w: 2, h: 0.5, margin: 0 },
    );
    slide.addTable(
      [[{ text: 'Table 0.1 inch', options: {} }]],
      { x: 0.2, y: 1.8, w: 2, h: 0.5, margin: 0.1 },
    );
    slide.addTable(
      [[{ text: 'Table inch tuple', options: {} }]],
      { x: 0.2, y: 2.6, w: 2, h: 0.5, margin: [0.05, 0.1, 0.15, 0.2] },
    );
    slide.addTable(
      [[
        { text: 'Inherited 0.1', options: {} },
        { text: 'Cell zero', options: { margin: 0 } },
        { text: 'Cell quarter inch', options: { margin: 0.25 } },
        { text: 'Cell inch tuple', options: { margin: [0.05, 0.1, 0.15, 0.2] } },
        { text: 'Cell scalar one point', options: { margin: 1 } },
        { text: 'Cell point tuple', options: { margin: [1, 2, 3, 4] } },
        { text: 'Cell negative inch', options: { margin: -0.1 } },
      ]],
      { x: 0.2, y: 3.4, w: 12.8, h: 1, margin: 0.1 },
    );

    const document = await importPptxGenJS(generated);
    const tables = document.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(5);
    const snapshots = tables.map((table) =>
      table.rows[0]!.cells.map(({ margins }) => margins));
    expect(snapshots).toEqual([
      [{ top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 }],
      [{ top: 0, right: 0, bottom: 0, left: 0 }],
      [{ top: 7.2, right: 7.2, bottom: 7.2, left: 7.2 }],
      [{ top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 }],
      [
        { top: 7.2, right: 7.2, bottom: 7.2, left: 7.2 },
        { top: 0, right: 0, bottom: 0, left: 0 },
        { top: 18, right: 18, bottom: 18, left: 18 },
        { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 },
        { top: 1, right: 1, bottom: 1, left: 1 },
        { top: 1, right: 2, bottom: 3, left: 4 },
        { top: -7.2, right: -7.2, bottom: -7.2, left: -7.2 },
      ],
    ]);
    expect(tables.flatMap(({ rows }) => rows[0]!.cells.map(({ text }) => text))).toEqual([
      'Omitted defaults',
      'Table zero',
      'Table 0.1 inch',
      'Table inch tuple',
      'Inherited 0.1',
      'Cell zero',
      'Cell quarter inch',
      'Cell inch tuple',
      'Cell scalar one point',
      'Cell point tuple',
      'Cell negative inch',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(xml.match(/<a:tcPr[^>]* marL=/g)).toHaveLength(11);
    expect(xml.match(/<a:tcPr[^>]* marR=/g)).toHaveLength(11);
    expect(xml.match(/<a:tcPr[^>]* marT=/g)).toHaveLength(11);
    expect(xml.match(/<a:tcPr[^>]* marB=/g)).toHaveLength(11);
    expect(xml.match(/<a:tcPr marL="91440" marR="91440" marT="91440" marB="91440">/g)).toHaveLength(2);
    expect(xml.match(/<a:tcPr marL="182880" marR="91440" marT="45720" marB="137160">/g)).toHaveLength(2);
    expect(xml).toContain('<a:tcPr marL="12700" marR="12700" marT="12700" marB="12700">');
    expect(xml).toContain('<a:tcPr marL="50800" marR="25400" marT="12700" marB="38100">');
    expect(xml).toContain('<a:tcPr marL="-91440" marR="-91440" marT="-91440" marB="-91440">');
    expect(xml).not.toMatch(/<a:bodyPr[^>]*(?:lIns|rIns|tIns|bIns)=/);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables.map((table) =>
      table.rows[0]!.cells.map(({ margins }) => margins))).toEqual(snapshots);
  });

  it('matches native table-cell margin creation to supported PptxGenJS output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable([[
      { text: 'Default', options: {} },
      { text: 'Zero', options: { margin: 0 } },
      { text: 'One point', options: { margin: 1 } },
      { text: 'Seven point two', options: { margin: 0.1 } },
      { text: 'TRBL', options: { margin: [0.05, 0.1, 0.15, 0.2] } },
      { text: 'Negative', options: { margin: -0.1 } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 10.5,
      h: 1,
      colW: [1.75, 1.75, 1.75, 1.75, 1.75, 1.75],
      rowH: 1,
    });
    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(importedTable).toBeInstanceOf(TableModel);

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeTable = native.addSlide().addTable([[
      { text: 'Default' },
      { text: 'Zero', options: { margin: 0 } },
      { text: 'One point', options: { margin: 1 } },
      { text: 'Seven point two', options: { margin: 7.2 } },
      { text: 'TRBL', options: { margin: [3.6, 7.2, 10.8, 14.4] } },
      { text: 'Negative', options: { margin: -7.2 } },
    ]], {
      x: inches(0.5),
      y: inches(0.5),
      columnWidths: inches(1.75),
      rowHeights: inches(1),
    });
    const expectedMargins = [
      { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 },
      { top: 0, right: 0, bottom: 0, left: 0 },
      { top: 1, right: 1, bottom: 1, left: 1 },
      { top: 7.2, right: 7.2, bottom: 7.2, left: 7.2 },
      { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 },
      { top: -7.2, right: -7.2, bottom: -7.2, left: -7.2 },
    ];
    const expectedText = [
      'Default',
      'Zero',
      'One point',
      'Seven point two',
      'TRBL',
      'Negative',
    ];
    expect(nativeTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(expectedMargins);
    expect(importedTable!.rows[0]!.cells.map(({ margins }) => margins)).toEqual(expectedMargins);
    expect(nativeTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(importedTable!.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(nativeTable.transform).toMatchObject(importedTable!.transform);
    expect(nativeTable.columnWidths).toEqual(importedTable!.columnWidths);
    expect(nativeTable.rowHeights).toEqual(importedTable!.rowHeights);

    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(native.slides[0]!.partUri).bytes,
    );
    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    const directTokens = [
      'marL="91440" marR="91440" marT="45720" marB="45720"',
      'marL="0" marR="0" marT="0" marB="0"',
      'marL="12700" marR="12700" marT="12700" marB="12700"',
      'marL="91440" marR="91440" marT="91440" marB="91440"',
      'marL="182880" marR="91440" marT="45720" marB="137160"',
      'marL="-91440" marR="-91440" marT="-91440" marB="-91440"',
    ];
    for (const xml of [nativeXml, importedXml]) {
      for (const token of directTokens) expect(xml).toContain(token);
    }

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows).toEqual(nativeTable.rows);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(importedTable!.rows);

    const nativeDifferences = PptxDocument.create({ slideSize: 'wide' });
    const nativeDifferenceTable = nativeDifferences.addSlide().addTable([[
      { text: 'Native 0.1 point', options: { margin: 0.1 } },
      { text: 'Native partial', options: { margin: { top: 2, left: -2 } } },
    ]]);
    expect(nativeDifferenceTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual([
      { top: 0.1, right: 0.1, bottom: 0.1, left: 0.1 },
      { top: 2, right: 7.2, bottom: 3.6, left: -2 },
    ]);
    const nativeDifferencesXml = new TextDecoder().decode(
      nativeDifferences.opcPackage.requirePart(nativeDifferences.slides[0]!.partUri).bytes,
    );
    expect(nativeDifferencesXml).toContain(
      '<a:tcPr marL="1270" marR="1270" marT="1270" marB="1270">',
    );
    expect(nativeDifferencesXml).toContain(
      '<a:tcPr marL="-25400" marR="91440" marT="25400" marB="45720">',
    );
    expect(importedXml).toContain(
      '<a:tcPr marL="91440" marR="91440" marT="91440" marB="91440">',
    );
  });

  it('matches native table-level margin creation to PptxGenJS final state', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable([[
      { text: 'Inherited string', options: {} },
      { text: 'Inherited object', options: {} },
      { text: 'Zero', options: {
        margin: 0,
        border: { type: 'solid', color: '336699', pt: 1 },
        fill: { color: 'DDEEFF' },
      } },
      { text: 'Points', options: {
        margin: [1, 2, 3, 4],
        border: { type: 'dash', color: 'CC3300', pt: 1.5 },
        fill: { color: '112233', transparency: 25 },
      } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 8,
      h: 1,
      colW: [2, 2, 2, 2],
      rowH: [1],
      margin: [0.05, 0.1, 0.15, 0.2],
      valign: 'middle',
    });
    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes[0] as TableModel;

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeSlide = native.addSlide();
    const nativeTable = nativeSlide.addTable([[
      'Inherited string',
      { text: 'Inherited object' },
      { text: 'Zero', options: {
        margin: 0,
        border: {
          kind: 'line',
          color: { kind: 'srgb', value: '336699' },
          width: 1,
          style: 'solid',
        },
        fill: { kind: 'solid', color: { kind: 'srgb', value: 'DDEEFF' } },
      } },
      { text: 'Points', options: {
        margin: [1, 2, 3, 4],
        border: {
          kind: 'line',
          color: { kind: 'srgb', value: 'CC3300' },
          width: 1.5,
          style: 'dash',
        },
        fill: {
          kind: 'solid',
          color: { kind: 'srgb', value: '112233' },
          transparency: 25,
        },
      } },
    ]], {
      x: inches(0.5),
      y: inches(0.5),
      columnWidths: inches(2),
      rowHeights: inches(1),
      margin: [3.6, 7.2, 10.8, 14.4],
      valign: 'middle',
    });
    const expectedMargins = [
      { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 },
      { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 },
      { top: 0, right: 0, bottom: 0, left: 0 },
      { top: 1, right: 2, bottom: 3, left: 4 },
    ];
    const expectedText = ['Inherited string', 'Inherited object', 'Zero', 'Points'];
    expect(nativeTable.rows[0]!.cells.map(({ margins }) => margins))
      .toEqual(expectedMargins);
    expect(importedTable.rows[0]!.cells.map(({ margins }) => margins))
      .toEqual(expectedMargins);
    expect(nativeTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(importedTable.rows[0]!.cells.map(({ text }) => text)).toEqual(expectedText);
    expect(nativeTable.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(Array(4).fill('middle'));
    expect(importedTable.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(Array(4).fill('middle'));
    expect(nativeTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(
      importedTable.rows[0]!.cells.map(({ borders }) => borders),
    );
    expect(nativeTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(
      importedTable.rows[0]!.cells.map(({ fill }) => fill),
    );
    expect(nativeTable.transform).toMatchObject(importedTable.transform);
    expect(nativeTable.columnWidths).toEqual(importedTable.columnWidths);
    expect(nativeTable.rowHeights).toEqual(importedTable.rowHeights);

    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(nativeSlide.partUri).bytes,
    );
    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    const directMarginVectors = (xml: string) => [...xml.matchAll(
      /<a:tcPr marL="(-?\d+)" marR="(-?\d+)" marT="(-?\d+)" marB="(-?\d+)"/g,
    )].map((match) => match.slice(1).map(Number));
    const expectedMarginVectors = [
      [182880, 91440, 45720, 137160],
      [182880, 91440, 45720, 137160],
      [0, 0, 0, 0],
      [50800, 25400, 12700, 38100],
    ];
    expect(directMarginVectors(nativeXml)).toEqual(expectedMarginVectors);
    expect(directMarginVectors(importedXml)).toEqual(expectedMarginVectors);
    for (const xml of [nativeXml, importedXml]) {
      expect([...xml.matchAll(/<a:tcPr[^>]* anchor="([^"]+)"/g)]
        .map((match) => match[1])).toEqual(Array(4).fill('ctr'));
      expect(xml).not.toMatch(/<a:bodyPr[^>]*(?:lIns|rIns|tIns|bIns|anchor)=/);
    }

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(nativeTable.rows);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(importedTable.rows);

    const generatedDifference = new PptxGenJS();
    generatedDifference.addSlide().addTable(
      [[{ text: 'PptxGenJS 0.1 table', options: {} }]],
      { x: 1, y: 1, w: 2, h: 1, colW: [2], rowH: [1], margin: 0.1 },
    );
    const importedDifference = await importPptxGenJS(generatedDifference);
    const nativeDifference = PptxDocument.create();
    const nativeDifferenceTable = nativeDifference.addSlide().addTable(
      [['Native 0.1 table']],
      { margin: 0.1 },
    );
    expect((importedDifference.slides[0]!.shapes[0] as TableModel)
      .rows[0]!.cells[0]!.margins).toEqual({
      top: 7.2,
      right: 7.2,
      bottom: 7.2,
      left: 7.2,
    });
    expect(nativeDifferenceTable.rows[0]!.cells[0]!.margins).toEqual({
      top: 0.1,
      right: 0.1,
      bottom: 0.1,
      left: 0.1,
    });

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    expect(() => nativeSlide.addTable([['Invalid']], { margin: null as never }))
      .toThrow(TypeError);
    expect(() => nativeSlide.addTable([['Invalid']], { margin: [1, 2, 3] as never }))
      .toThrow(RangeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
  });

  it('matches native table-level solid fill creation to PptxGenJS final state', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const generatedSlide = generated.addSlide();
    generatedSlide.addTable([[
      { text: 'Inherited string', options: {} },
      { text: 'Inherited object', options: {} },
      { text: 'Cell yellow', options: {
        fill: { color: 'FFFF00', transparency: 50 },
        border: { type: 'solid', color: '336699', pt: 1 },
      } },
      { text: 'Cell full transparency', options: {
        fill: { color: '445566', transparency: 100 },
      } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 8,
      h: 1,
      colW: [2, 2, 2, 2],
      rowH: [1],
      fill: { color: generated.SchemeColor.accent1, transparency: 25 },
      margin: 0.1,
      valign: 'middle',
    });
    generatedSlide.addTable(
      [[{ text: 'Table sRGB opaque', options: {} }]],
      {
        x: 0.5,
        y: 2,
        w: 2,
        h: 0.5,
        colW: [2],
        rowH: [0.5],
        fill: { color: 'FF0000' },
      },
    );
    generatedSlide.addTable(
      [[{ text: 'Table fractional', options: {} }]],
      {
        x: 0.5,
        y: 2.8,
        w: 2,
        h: 0.5,
        colW: [2],
        rowH: [0.5],
        fill: { color: '112233', transparency: 33.333 },
      },
    );
    generatedSlide.addTable(
      [[{ text: 'Table full transparency', options: {} }]],
      {
        x: 0.5,
        y: 3.6,
        w: 2,
        h: 0.5,
        colW: [2],
        rowH: [0.5],
        fill: { color: '445566', transparency: 100 },
      },
    );
    const imported = await importPptxGenJS(generated);
    const importedTables = imported.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeSlide = native.addSlide();
    const nativeTable = nativeSlide.addTable([[
      'Inherited string',
      { text: 'Inherited object' },
      { text: 'Cell yellow', options: {
        fill: {
          kind: 'solid',
          color: { kind: 'srgb', value: 'FFFF00' },
          transparency: 50,
        },
        border: {
          kind: 'line',
          color: { kind: 'srgb', value: '336699' },
          width: 1,
          style: 'solid',
        },
      } },
      { text: 'Cell full transparency', options: { fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: '445566' },
        transparency: 100,
      } } },
    ]], {
      x: inches(0.5),
      y: inches(0.5),
      width: inches(8),
      height: inches(1),
      columnWidths: inches(2),
      rowHeights: inches(1),
      fill: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 25,
      },
      margin: 7.2,
      valign: 'middle',
    });
    nativeSlide.addTable([['Table sRGB opaque']], {
      x: inches(0.5),
      y: inches(2),
      width: inches(2),
      height: inches(0.5),
      columnWidths: [inches(2)],
      rowHeights: [inches(0.5)],
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
      },
    });
    nativeSlide.addTable([['Table fractional']], {
      x: inches(0.5),
      y: inches(2.8),
      width: inches(2),
      height: inches(0.5),
      columnWidths: [inches(2)],
      rowHeights: [inches(0.5)],
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: '112233' },
        transparency: 33.333,
      },
    });
    nativeSlide.addTable([['Table full transparency']], {
      x: inches(0.5),
      y: inches(3.6),
      width: inches(2),
      height: inches(0.5),
      columnWidths: [inches(2)],
      rowHeights: [inches(0.5)],
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: '445566' },
        transparency: 100,
      },
    });
    const nativeTables = nativeSlide.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );

    const expectedFills = [
      [
        {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent1' },
          transparency: 25,
        },
        {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent1' },
          transparency: 25,
        },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: 'FFFF00' },
          transparency: 50,
        },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: '445566' },
          transparency: 100,
        },
      ],
      [{ kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } }],
      [{
        kind: 'solid',
        color: { kind: 'srgb', value: '112233' },
        transparency: 33.333,
      }],
      [{
        kind: 'solid',
        color: { kind: 'srgb', value: '445566' },
        transparency: 100,
      }],
    ];
    expect(nativeTables.map((table) =>
      table.rows[0]!.cells.map(({ fill }) => fill))).toEqual(expectedFills);
    expect(importedTables.map((table) =>
      table.rows[0]!.cells.map(({ fill }) => fill))).toEqual(expectedFills);
    expect(nativeTable.rows[0]!.cells.map(({ text }) => text)).toEqual(
      importedTables[0]!.rows[0]!.cells.map(({ text }) => text),
    );
    expect(nativeTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(
      importedTables[0]!.rows[0]!.cells.map(({ margins }) => margins),
    );
    expect(nativeTable.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(
      importedTables[0]!.rows[0]!.cells.map(
        ({ verticalAlignment }) => verticalAlignment),
    );
    expect(nativeTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(
      importedTables[0]!.rows[0]!.cells.map(({ borders }) => borders),
    );
    for (let index = 0; index < nativeTables.length; index += 1) {
      expect(nativeTables[index]!.transform).toMatchObject(
        importedTables[index]!.transform,
      );
      expect(nativeTables[index]!.columnWidths).toEqual(
        importedTables[index]!.columnWidths,
      );
      expect(nativeTables[index]!.rowHeights).toEqual(
        importedTables[index]!.rowHeights,
      );
    }

    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(nativeSlide.partUri).bytes,
    );
    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    for (const xml of [nativeXml, importedXml]) {
      expect(xml.match(
        /<a:solidFill><a:schemeClr val="accent1"><a:alpha val="75000"\/><\/a:schemeClr><\/a:solidFill>/g,
      )).toHaveLength(2);
      expect(xml).toContain(
        '<a:solidFill><a:srgbClr val="FFFF00">'
        + '<a:alpha val="50000"/></a:srgbClr></a:solidFill>',
      );
      expect(xml).toContain(
        '<a:solidFill><a:srgbClr val="112233">'
        + '<a:alpha val="66667"/></a:srgbClr></a:solidFill>',
      );
      expect(xml.match(
        /<a:solidFill><a:srgbClr val="445566"><a:alpha val="0"\/><\/a:srgbClr><\/a:solidFill>/g,
      )).toHaveLength(2);
      for (const properties of xml.matchAll(
        /<a:tcPr[^>]*>([\s\S]*?)<\/a:tcPr>/g,
      )) {
        const direct = properties[1]!;
        const fill = Math.max(
          direct.lastIndexOf('<a:noFill/>'),
          direct.lastIndexOf('<a:solidFill>'),
        );
        expect(fill).toBeGreaterThan(direct.indexOf('</a:lnB>'));
      }
    }

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect((reopenedNative.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    )).map((table) => table.rows[0]!.cells.map(({ fill }) => fill)))
      .toEqual(expectedFills);
    expect((reopenedImported.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    )).map((table) => table.rows[0]!.cells.map(({ fill }) => fill)))
      .toEqual(expectedFills);

    const generatedDifferences = new PptxGenJS();
    generatedDifferences.addSlide().addTable([[
      { text: 'Collapsed none', options: {} },
      { text: 'Collapsed zero alpha', options: {
        fill: { color: '00FF00', transparency: 0 },
      } },
    ]], {
      x: 1,
      y: 1,
      w: 4,
      h: 1,
      colW: [2, 2],
      rowH: [1],
      fill: { type: 'none' },
    });
    const importedDifferences = await importPptxGenJS(generatedDifferences);
    const importedDifferenceTable =
      importedDifferences.slides[0]!.shapes[0] as TableModel;
    expect(importedDifferenceTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual([
      undefined,
      { kind: 'solid', color: { kind: 'srgb', value: '00FF00' } },
    ]);

    const nativeDifferences = PptxDocument.create();
    const nativeDifferenceSlide = nativeDifferences.addSlide();
    const nativeDifferenceTable = nativeDifferenceSlide.addTable([[
      'Inherited none',
      { text: 'Explicit zero override', options: { fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: '00FF00' },
        transparency: 0,
      } } },
    ]], { fill: { kind: 'none' } });
    expect(nativeDifferenceTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual([
      { kind: 'none' },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: '00FF00' },
        transparency: 0,
      },
    ]);
    const nativeDifferenceXml = new TextDecoder().decode(
      nativeDifferences.opcPackage.requirePart(nativeDifferenceSlide.partUri).bytes,
    );
    expect(nativeDifferenceXml).toContain('</a:lnB><a:noFill/></a:tcPr>');
    expect(nativeDifferenceXml).toContain(
      '<a:solidFill><a:srgbClr val="00FF00">'
      + '<a:alpha val="100000"/></a:srgbClr></a:solidFill>',
    );
    const reopenedNativeDifference = await PptxDocument.open(
      await nativeDifferences.write(),
    );
    expect((reopenedNativeDifference.slides[0]!.shapes[0] as TableModel)
      .rows[0]!.cells.map(({ fill }) => fill)).toEqual([
      { kind: 'none' },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: '00FF00' },
        transparency: 0,
      },
    ]);
    const reopenedImportedDifference = await PptxDocument.open(
      await importedDifferences.write(),
    );
    expect((reopenedImportedDifference.slides[0]!.shapes[0] as TableModel)
      .rows[0]!.cells.map(({ fill }) => fill)).toEqual([
      undefined,
      { kind: 'solid', color: { kind: 'srgb', value: '00FF00' } },
    ]);

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    expect(() => nativeSlide.addTable([['Invalid']], { fill: {} as never }))
      .toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
    expect(nativeSlide.shapes).toHaveLength(4);
  });

  it('imports PptxGenJS table-cell fills from direct cell properties', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    slide.addTable(
      [[{ text: 'Omitted fill', options: {} }]],
      { x: 0.2, y: 0.2, w: 2, h: 0.5 },
    );
    slide.addTable(
      [[{ text: 'Type none', options: { fill: { type: 'none' } } }]],
      { x: 0.2, y: 1, w: 2, h: 0.5 },
    );
    slide.addTable(
      [[{ text: 'Table red', options: {} }]],
      { x: 0.2, y: 1.8, w: 2, h: 0.5, fill: { color: 'FF0000' } },
    );
    slide.addTable(
      [[{ text: 'Table theme alpha', options: {} }]],
      {
        x: 0.2,
        y: 2.6,
        w: 2,
        h: 0.5,
        fill: { color: generated.SchemeColor.accent1, transparency: 25 },
      },
    );
    slide.addTable(
      [[
        { text: 'Inherited blue', options: {} },
        { text: 'Cell yellow alpha', options: { fill: { color: 'FFFF00', transparency: 50 } } },
        { text: 'Explicit zero', options: { fill: { color: '00FF00', transparency: 0 } } },
        { text: 'Fractional', options: { fill: { color: '112233', transparency: 33.333 } } },
        { text: 'Full transparency', options: { fill: { color: '445566', transparency: 100 } } },
        { text: 'Deprecated alpha', options: { fill: { color: generated.SchemeColor.accent2, alpha: 25 } } },
        { text: 'Runtime negative', options: { fill: { color: '778899', transparency: -1 } } },
        { text: 'Runtime overflow', options: { fill: { color: 'AABBCC', transparency: 101 } } },
      ]],
      { x: 0.2, y: 3.4, w: 12.8, h: 1, fill: { color: '0000FF' } },
    );

    const document = await importPptxGenJS(generated);
    const tables = document.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(5);
    const snapshots = tables.map((table) => table.rows[0]!.cells.map(({ fill }) => fill));
    expect(snapshots).toEqual([
      [undefined],
      [undefined],
      [{ kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } }],
      [{
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 25,
      }],
      [
        { kind: 'solid', color: { kind: 'srgb', value: '0000FF' } },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: 'FFFF00' },
          transparency: 50,
        },
        { kind: 'solid', color: { kind: 'srgb', value: '00FF00' } },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: '112233' },
          transparency: 33.333,
        },
        {
          kind: 'solid',
          color: { kind: 'srgb', value: '445566' },
          transparency: 100,
        },
        {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 25,
        },
        undefined,
        undefined,
      ],
    ]);
    expect(tables.flatMap(({ rows }) => rows[0]!.cells.map(({ text }) => text))).toEqual([
      'Omitted fill',
      'Type none',
      'Table red',
      'Table theme alpha',
      'Inherited blue',
      'Cell yellow alpha',
      'Explicit zero',
      'Fractional',
      'Full transparency',
      'Deprecated alpha',
      'Runtime negative',
      'Runtime overflow',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    const directFillXml = [...xml.matchAll(/<a:tcPr[^>]*>([\s\S]*?)<\/a:tcPr>/g)]
      .map((match) => match[1]!.replace(/<a:ln[LRBT][\s\S]*?<\/a:ln[LRBT]>/g, ''))
      .map((properties) => properties.match(/<a:solidFill>[\s\S]*?<\/a:solidFill>/)?.[0]);
    expect(directFillXml.filter(Boolean)).toHaveLength(10);
    expect(directFillXml[0]).toBeUndefined();
    expect(directFillXml[1]).toBeUndefined();
    expect(xml).toContain('<a:solidFill><a:schemeClr val="accent1"><a:alpha val="75000"/></a:schemeClr></a:solidFill>');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="112233"><a:alpha val="66667"/></a:srgbClr></a:solidFill>');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="445566"><a:alpha val="0"/></a:srgbClr></a:solidFill>');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="778899"><a:alpha val="101000"/></a:srgbClr></a:solidFill>');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="AABBCC"><a:alpha val="-1000"/></a:srgbClr></a:solidFill>');

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables.map((table) =>
      table.rows[0]!.cells.map(({ fill }) => fill))).toEqual(snapshots);

    const conformanceGenerated = new PptxGenJS();
    conformanceGenerated.layout = 'LAYOUT_WIDE';
    conformanceGenerated.addSlide().addTable(
      [[
        { text: 'Opaque', options: { fill: { color: 'FF0000' } } },
        { text: 'Theme alpha', options: {
          fill: { color: conformanceGenerated.SchemeColor.accent1, transparency: 25 },
        } },
      ]],
      { x: 1, y: 1, w: 4, h: 1, colW: [2, 2], rowH: [1] },
    );
    const conformanceImported = await importPptxGenJS(conformanceGenerated);
    const importedConformanceTable = conformanceImported.slides[0]!.shapes[0] as TableModel;
    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeTable = native.addSlide().addTable(
      [[
        { text: 'Opaque', options: { fill: {
          kind: 'solid',
          color: { kind: 'srgb', value: 'FF0000' },
        } } },
        { text: 'Theme alpha', options: { fill: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent1' },
          transparency: 25,
        } } },
      ]],
      {
        x: inches(1),
        y: inches(1),
        width: inches(4),
        height: inches(1),
        columnWidths: [inches(2), inches(2)],
        rowHeights: [inches(1)],
      },
    );
    expect(nativeTable.transform).toMatchObject(importedConformanceTable.transform);
    expect(nativeTable.columnWidths).toEqual(importedConformanceTable.columnWidths);
    expect(nativeTable.rowHeights).toEqual(importedConformanceTable.rowHeights);
    expect(nativeTable.rows.map(({ cells }) => cells.map(({ fill }) => fill))).toEqual(
      importedConformanceTable.rows.map(({ cells }) => cells.map(({ fill }) => fill)),
    );
    expect(nativeTable.rows.map(({ cells }) => cells.map(({ margins }) => margins))).toEqual(
      importedConformanceTable.rows.map(({ cells }) => cells.map(({ margins }) => margins)),
    );
    expect(nativeTable.rows.map(({ cells }) => cells.map(({ borders }) => borders))).toEqual(
      importedConformanceTable.rows.map(({ cells }) => cells.map(({ borders }) => borders)),
    );
    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(native.slides[0]!.partUri).bytes,
    );
    expect(nativeXml).toContain(
      '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>',
    );
    expect(nativeXml).toContain(
      '<a:solidFill><a:schemeClr val="accent1"><a:alpha val="75000"/></a:schemeClr></a:solidFill>',
    );
    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedConformance = await PptxDocument.open(await conformanceImported.write());
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows).toEqual(nativeTable.rows);
    expect((reopenedConformance.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(importedConformanceTable.rows);

    const nativeDifferences = PptxDocument.create({ slideSize: 'wide' });
    nativeDifferences.addSlide().addTable([[
      { text: 'Explicit zero', options: { fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: '00FF00' },
        transparency: 0,
      } } },
      { text: 'Direct none', options: { fill: { kind: 'none' } } },
    ]]);
    const nativeDifferencesXml = new TextDecoder().decode(
      nativeDifferences.opcPackage.requirePart(nativeDifferences.slides[0]!.partUri).bytes,
    );
    expect(nativeDifferencesXml).toContain(
      '<a:solidFill><a:srgbClr val="00FF00"><a:alpha val="100000"/></a:srgbClr></a:solidFill>',
    );
    expect(nativeDifferencesXml).toContain('</a:lnB><a:noFill/></a:tcPr>');
    expect(xml).toContain('<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>');
    expect(directFillXml[1]).toBeUndefined();
  });

  it('imports PptxGenJS table-cell borders from materialized direct lines', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const slide = generated.addSlide();
    const tableTuple = [
      { type: 'none' },
      { type: 'dash', color: '00FF00', pt: 1.5 },
      { type: 'solid', color: '0000FF', pt: 0 },
      { type: 'solid' },
    ] satisfies [BorderProps, BorderProps, BorderProps, BorderProps];
    const partialCellTuple = [
      undefined,
      { type: 'dash' },
      undefined,
      { type: 'solid' },
    ] as unknown as [BorderProps, BorderProps, BorderProps, BorderProps];

    slide.addTable(
      [[{ text: 'Omitted border', options: {} }]],
      { x: 0.2, y: 0.2, w: 2, h: 0.5 },
    );
    slide.addTable(
      [[{ text: 'Table scalar', options: {} }]],
      { x: 0.2, y: 1, w: 2, h: 0.5, border: { type: 'solid', color: 'FF0000', pt: 2 } },
    );
    slide.addTable(
      [[{ text: 'Table tuple', options: {} }]],
      { x: 0.2, y: 1.8, w: 2, h: 0.5, border: tableTuple },
    );
    slide.addTable(
      [[{
        text: 'Cell scalar zero',
        options: { border: { type: 'solid', color: 'FFFF00', pt: 0 } },
      }]],
      { x: 0.2, y: 2.6, w: 2, h: 0.5, border: { type: 'solid', color: 'AAAAAA', pt: 3 } },
    );
    slide.addTable(
      [[{ text: 'Cell partial tuple', options: { border: partialCellTuple } }]],
      { x: 0.2, y: 3.4, w: 2, h: 0.5 },
    );
    slide.addTable(
      [[{ text: 'Default border values', options: {} }]],
      { x: 0.2, y: 4.2, w: 2, h: 0.5, border: {} },
    );
    slide.addTable(
      [[
        { text: 'Runtime negative', options: { border: { type: 'solid', color: '778899', pt: -1 } } },
        { text: 'Runtime overflow', options: { border: { type: 'dash', color: 'AABBCC', pt: 2000 } } },
      ]],
      { x: 0.2, y: 5, w: 4, h: 0.5 },
    );

    const document = await importPptxGenJS(generated);
    const tables = document.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(tables).toHaveLength(7);
    const line = (
      color: string,
      width: number,
      style: 'solid' | 'dash' = 'solid',
    ) => ({
      kind: 'line' as const,
      color: { kind: 'srgb' as const, value: color },
      width,
      style,
    });
    const none = { kind: 'none' as const };
    const four = <T>(value: T) => ({ top: value, right: value, bottom: value, left: value });
    const snapshots = tables.map((table) =>
      table.rows[0]!.cells.map(({ borders }) => borders));
    expect(snapshots).toEqual([
      [four(none)],
      [four(line('FF0000', 2))],
      [{
        top: none,
        right: line('00FF00', 1.5, 'dash'),
        bottom: line('0000FF', 1),
        left: line('666666', 1),
      }],
      [four(line('FFFF00', 0))],
      [{
        top: none,
        right: line('666666', 1, 'dash'),
        bottom: none,
        left: line('666666', 1),
      }],
      [four(line('666666', 1))],
      [undefined, undefined],
    ]);
    expect(tables.flatMap(({ rows }) => rows[0]!.cells.map(({ text }) => text))).toEqual([
      'Omitted border',
      'Table scalar',
      'Table tuple',
      'Cell scalar zero',
      'Cell partial tuple',
      'Default border values',
      'Runtime negative',
      'Runtime overflow',
    ]);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(xml).toContain(
      '<a:lnL w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnL><a:lnR w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnR><a:lnT w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnT><a:lnB w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnB>',
    );
    expect(xml).toContain('<a:prstDash val="solid"/>');
    expect(xml).toContain('<a:prstDash val="sysDash"/>');
    expect(xml).toContain('<a:lnB w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>');
    expect(xml).toContain('<a:lnL w="-12700" cap="flat" cmpd="sng" algn="ctr">');
    expect(xml).toContain('<a:lnL w="25400000" cap="flat" cmpd="sng" algn="ctr">');
    for (const properties of xml.matchAll(/<a:tcPr[^>]*>([\s\S]*?)<\/a:tcPr>/g)) {
      const direct = properties[1]!;
      const left = direct.indexOf('<a:lnL ');
      const right = direct.indexOf('<a:lnR ');
      const top = direct.indexOf('<a:lnT ');
      const bottom = direct.indexOf('<a:lnB ');
      expect(left).toBeGreaterThanOrEqual(0);
      expect(left).toBeLessThan(right);
      expect(right).toBeLessThan(top);
      expect(top).toBeLessThan(bottom);
    }

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTables = reopened.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(reopenedTables.map((table) =>
      table.rows[0]!.cells.map(({ borders }) => borders))).toEqual(snapshots);
  });

  it('matches native table-cell border creation to supported PptxGenJS output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    generated.addSlide().addTable(
      [[
        { text: 'Scalar', options: { border: {
          type: 'solid',
          color: 'FF0000',
          pt: 2,
        } } },
        { text: 'Tuple', options: { border: [
          { type: 'none' },
          { type: 'dash', color: '00FF00', pt: 1.5 },
          { type: 'solid', color: '0000FF', pt: 0 },
          { type: 'solid', color: '666666', pt: 1 },
        ] } },
      ]],
      { x: 1, y: 1, w: 4, h: 1, colW: [2, 2], rowH: [1] },
    );
    const imported = await importPptxGenJS(generated);
    const importedTable = imported.slides[0]!.shapes[0] as TableModel;

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeTable = native.addSlide().addTable(
      [[
        { text: 'Scalar', options: { border: {
          kind: 'line',
          color: { kind: 'srgb', value: 'FF0000' },
          width: 2,
          style: 'solid',
        } } },
        { text: 'Tuple', options: { border: [
          { kind: 'none' },
          {
            kind: 'line',
            color: { kind: 'srgb', value: '00FF00' },
            width: 1.5,
            style: 'dash',
          },
          {
            kind: 'line',
            color: { kind: 'srgb', value: '0000FF' },
            width: 0,
            style: 'solid',
          },
          {
            kind: 'line',
            color: { kind: 'srgb', value: '666666' },
            width: 1,
            style: 'solid',
          },
        ] } },
      ]],
      {
        x: inches(1),
        y: inches(1),
        width: inches(4),
        height: inches(1),
        columnWidths: [inches(2), inches(2)],
        rowHeights: [inches(1)],
      },
    );

    const line = (
      color: string,
      width: number,
      style: 'solid' | 'dash' = 'solid',
    ) => ({
      kind: 'line' as const,
      color: { kind: 'srgb' as const, value: color },
      width,
      style,
    });
    const none = { kind: 'none' as const };
    const four = <T>(value: T) => ({ top: value, right: value, bottom: value, left: value });
    const expectedBorders = [
      four(line('FF0000', 2)),
      {
        top: none,
        right: line('00FF00', 1.5, 'dash'),
        bottom: line('0000FF', 0),
        left: line('666666', 1),
      },
    ];
    expect(nativeTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(expectedBorders);
    expect(importedTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(expectedBorders);
    expect(nativeTable.rows[0]!.cells.map(({ margins }) => margins)).toEqual(
      importedTable.rows[0]!.cells.map(({ margins }) => margins),
    );
    expect(nativeTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(
      importedTable.rows[0]!.cells.map(({ fill }) => fill),
    );
    expect(nativeTable.transform).toMatchObject(importedTable.transform);
    expect(nativeTable.columnWidths).toEqual(importedTable.columnWidths);
    expect(nativeTable.rowHeights).toEqual(importedTable.rowHeights);

    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(native.slides[0]!.partUri).bytes,
    );
    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    for (const tableXml of [nativeXml, importedXml]) {
      for (const properties of tableXml.matchAll(/<a:tcPr[^>]*>([\s\S]*?)<\/a:tcPr>/g)) {
        const direct = properties[1]!;
        const left = direct.indexOf('<a:lnL ');
        const right = direct.indexOf('<a:lnR ');
        const top = direct.indexOf('<a:lnT ');
        const bottom = direct.indexOf('<a:lnB ');
        expect(left).toBeGreaterThanOrEqual(0);
        expect(left).toBeLessThan(right);
        expect(right).toBeLessThan(top);
        expect(top).toBeLessThan(bottom);
      }
      expect(tableXml).toContain('<a:lnR w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:prstDash val="sysDash"/>');
      expect(tableXml).toContain('<a:lnB w="0" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:prstDash val="solid"/>');
    }
    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect((reopenedNative.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(nativeTable.rows);
    expect((reopenedImported.slides[0]!.shapes[0] as TableModel).rows)
      .toEqual(importedTable.rows);

    const generatedDifferences = new PptxGenJS();
    generatedDifferences.addSlide().addTable(
      [[
        { text: 'PptxGenJS empty', options: { border: {} } },
        { text: 'PptxGenJS omitted type', options: { border: {
          color: '112233',
          pt: 2,
        } } },
      ]],
      { x: 1, y: 1, w: 4, h: 1, colW: [2, 2], rowH: [1] },
    );
    const importedDifferences = await importPptxGenJS(generatedDifferences);
    const importedDifferenceTable = importedDifferences.slides[0]!.shapes[0] as TableModel;
    expect(importedDifferenceTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual([
      four(line('666666', 1)),
      four(line('112233', 2)),
    ]);

    const nativeDifferences = PptxDocument.create();
    const nativeDifferenceTable = nativeDifferences.addSlide().addTable([[
      { text: 'Native empty', options: { border: {} } },
      { text: 'Native omitted style', options: { border: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
        width: 2,
      } } },
      { text: 'Native named theme', options: { border: {
        top: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent1' },
          width: 1,
          style: 'dash',
        },
        left: { kind: 'none' },
      } } },
      { text: 'Native zero', options: { border: {
        kind: 'line',
        color: { kind: 'srgb', value: '0000FF' },
        width: 0,
        style: 'solid',
      } } },
    ]]);
    expect(nativeDifferenceTable.rows[0]!.cells[0]!.borders).toEqual(four(none));
    const omittedStyle = {
      kind: 'line' as const,
      color: { kind: 'srgb' as const, value: '112233' },
      width: 2,
    };
    expect(nativeDifferenceTable.rows[0]!.cells[1]!.borders).toEqual(four(omittedStyle));
    expect(nativeDifferenceTable.rows[0]!.cells[2]!.borders).toEqual({
      top: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent1' },
        width: 1,
        style: 'dash',
      },
      right: none,
      bottom: none,
      left: none,
    });
    expect(nativeDifferenceTable.rows[0]!.cells[3]!.borders).toEqual(
      four(line('0000FF', 0)),
    );
    const nativeDifferencesXml = new TextDecoder().decode(
      nativeDifferences.opcPackage.requirePart(nativeDifferences.slides[0]!.partUri).bytes,
    );
    const cells = [...nativeDifferencesXml.matchAll(/<a:tc>([\s\S]*?)<\/a:tc>/g)]
      .map((match) => match[1]!);
    expect(cells[0]!.match(/<a:noFill\/>/g)).toHaveLength(4);
    expect(cells[1]).not.toContain('<a:prstDash');
    expect(cells[2]).toContain('<a:schemeClr val="accent1"/>');
    expect(cells[2]).toContain('<a:prstDash val="sysDash"/>');
    expect(cells[3]).toContain('<a:lnL w="0" cap="flat" cmpd="sng" algn="ctr">');
  });

  it('matches native table-level border creation to PptxGenJS final state', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const generatedSlide = generated.addSlide();
    generatedSlide.addTable([[
      { text: 'Inherited string', options: {} },
      { text: 'Inherited object', options: {} },
      { text: 'Cell partial', options: { border: [
        { type: 'none' },
        { type: 'none' },
        { type: 'dash', color: '70AD47', pt: 3 },
        { type: 'none' },
      ] } },
      { text: 'Cell none', options: { border: { type: 'none' } } },
    ]], {
      x: 0.5,
      y: 0.5,
      w: 12,
      h: 1,
      colW: [3, 3, 3, 3],
      rowH: [1],
      border: { type: 'dash', color: '4472C4', pt: 1.5 },
      fill: { color: 'D9EAF7' },
      margin: 0.1,
      valign: 'middle',
    });
    generatedSlide.addTable([[{ text: 'TRBL', options: {} }]], {
      x: 0.5,
      y: 2,
      w: 2,
      h: 0.5,
      colW: [2],
      rowH: [0.5],
      border: [
        { type: 'solid', color: 'FF0000', pt: 1 },
        { type: 'none' },
        { type: 'dash', color: '70AD47', pt: 2 },
        { type: 'solid', color: '4472C4', pt: 3 },
      ],
    });
    generatedSlide.addTable([[{ text: 'Scalar zero', options: {} }]], {
      x: 0.5,
      y: 3,
      w: 2,
      h: 0.5,
      colW: [2],
      rowH: [0.5],
      border: { type: 'solid', color: 'FF0000', pt: 0 },
    });
    generatedSlide.addTable([[{ text: 'Scalar none', options: {} }]], {
      x: 0.5,
      y: 4,
      w: 2,
      h: 0.5,
      colW: [2],
      rowH: [0.5],
      border: { type: 'none' },
    });
    const imported = await importPptxGenJS(generated);
    const importedTables = imported.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );

    const native = PptxDocument.create({ slideSize: 'wide' });
    const nativeSlide = native.addSlide();
    nativeSlide.addTable([[
      'Inherited string',
      { text: 'Inherited object', options: {} },
      { text: 'Cell partial', options: { border: {
        bottom: {
          kind: 'line',
          color: { kind: 'srgb', value: '70AD47' },
          width: 3,
          style: 'dash',
        },
      } } },
      { text: 'Cell none', options: { border: { kind: 'none' } } },
    ]], {
      x: inches(0.5),
      y: inches(0.5),
      width: inches(12),
      height: inches(1),
      columnWidths: inches(3),
      rowHeights: inches(1),
      border: {
        kind: 'line',
        color: { kind: 'srgb', value: '4472C4' },
        width: 1.5,
        style: 'dash',
      },
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'D9EAF7' },
      },
      margin: 7.2,
      valign: 'middle',
    });
    nativeSlide.addTable([['TRBL']], {
      x: inches(0.5),
      y: inches(2),
      width: inches(2),
      height: inches(0.5),
      columnWidths: [inches(2)],
      rowHeights: [inches(0.5)],
      border: [
        {
          kind: 'line',
          color: { kind: 'srgb', value: 'FF0000' },
          width: 1,
          style: 'solid',
        },
        { kind: 'none' },
        {
          kind: 'line',
          color: { kind: 'srgb', value: '70AD47' },
          width: 2,
          style: 'dash',
        },
        {
          kind: 'line',
          color: { kind: 'srgb', value: '4472C4' },
          width: 3,
          style: 'solid',
        },
      ],
    });
    nativeSlide.addTable([['Scalar zero']], {
      x: inches(0.5),
      y: inches(3),
      width: inches(2),
      height: inches(0.5),
      columnWidths: [inches(2)],
      rowHeights: [inches(0.5)],
      border: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        width: 0,
        style: 'solid',
      },
    });
    nativeSlide.addTable([['Scalar none']], {
      x: inches(0.5),
      y: inches(4),
      width: inches(2),
      height: inches(0.5),
      columnWidths: [inches(2)],
      rowHeights: [inches(0.5)],
      border: { kind: 'none' },
    });
    const nativeTables = nativeSlide.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );

    const line = (
      color: string,
      width: number,
      style: 'solid' | 'dash' = 'solid',
    ) => ({
      kind: 'line' as const,
      color: { kind: 'srgb' as const, value: color },
      width,
      style,
    });
    const none = { kind: 'none' as const };
    const four = <T>(value: T) => ({
      top: value,
      right: value,
      bottom: value,
      left: value,
    });
    const inherited = four(line('4472C4', 1.5, 'dash'));
    const expectedBorders = [
      [
        inherited,
        inherited,
        {
          top: none,
          right: none,
          bottom: line('70AD47', 3, 'dash'),
          left: none,
        },
        four(none),
      ],
      [{
        top: line('FF0000', 1),
        right: none,
        bottom: line('70AD47', 2, 'dash'),
        left: line('4472C4', 3),
      }],
      [four(line('FF0000', 0))],
      [four(none)],
    ];
    const borderMatrices = (tables: readonly TableModel[]) => tables.map((table) =>
      table.rows[0]!.cells.map(({ borders }) => borders));
    expect(borderMatrices(nativeTables)).toEqual(expectedBorders);
    expect(borderMatrices(importedTables)).toEqual(expectedBorders);
    expect(nativeTables[0]!.rows[0]!.cells.map(({ text }) => text)).toEqual(
      importedTables[0]!.rows[0]!.cells.map(({ text }) => text),
    );
    expect(nativeTables[0]!.rows[0]!.cells.map(({ fill }) => fill)).toEqual(
      importedTables[0]!.rows[0]!.cells.map(({ fill }) => fill),
    );
    expect(nativeTables[0]!.rows[0]!.cells.map(({ margins }) => margins)).toEqual(
      importedTables[0]!.rows[0]!.cells.map(({ margins }) => margins),
    );
    expect(nativeTables[0]!.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual(
      importedTables[0]!.rows[0]!.cells.map(
        ({ verticalAlignment }) => verticalAlignment),
    );
    for (let index = 0; index < nativeTables.length; index += 1) {
      expect(nativeTables[index]!.transform).toMatchObject(
        importedTables[index]!.transform,
      );
      expect(nativeTables[index]!.columnWidths).toEqual(
        importedTables[index]!.columnWidths,
      );
      expect(nativeTables[index]!.rowHeights).toEqual(
        importedTables[index]!.rowHeights,
      );
    }

    const nativeXml = new TextDecoder().decode(
      native.opcPackage.requirePart(nativeSlide.partUri).bytes,
    );
    const importedXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(imported.slides[0]!.partUri).bytes,
    );
    for (const tableXml of [nativeXml, importedXml]) {
      for (const properties of tableXml.matchAll(
        /<a:tcPr[^>]*>([\s\S]*?)<\/a:tcPr>/g,
      )) {
        const direct = properties[1]!;
        const left = direct.indexOf('<a:lnL ');
        const right = direct.indexOf('<a:lnR ');
        const top = direct.indexOf('<a:lnT ');
        const bottom = direct.indexOf('<a:lnB ');
        expect(left).toBeGreaterThanOrEqual(0);
        expect(left).toBeLessThan(right);
        expect(right).toBeLessThan(top);
        expect(top).toBeLessThan(bottom);
      }
      expect(tableXml).toContain(
        '<a:lnB w="38100" cap="flat" cmpd="sng" algn="ctr">'
        + '<a:solidFill><a:srgbClr val="70AD47"/></a:solidFill>'
        + '<a:prstDash val="sysDash"/>',
      );
    }

    const reopenedNative = await PptxDocument.open(await native.write());
    const reopenedImported = await PptxDocument.open(await imported.write());
    expect(borderMatrices(reopenedNative.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    ))).toEqual(expectedBorders);
    expect(borderMatrices(reopenedImported.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    ))).toEqual(expectedBorders);

    const generatedDifferences = new PptxGenJS();
    const generatedDifferenceSlide = generatedDifferences.addSlide();
    generatedDifferenceSlide.addTable(
      [[{ text: 'PptxGenJS empty table', options: {} }]],
      { x: 0.5, y: 0.5, w: 2, h: 0.5, border: {} },
    );
    generatedDifferenceSlide.addTable(
      [[{ text: 'PptxGenJS empty cell', options: { border: {} } }]],
      {
        x: 0.5,
        y: 1.5,
        w: 2,
        h: 0.5,
        border: { type: 'solid', color: 'FF0000', pt: 2 },
      },
    );
    generatedDifferenceSlide.addTable(
      [[{ text: 'PptxGenJS short tuple', options: {} }]],
      {
        x: 0.5,
        y: 2.5,
        w: 2,
        h: 0.5,
        border: [{ type: 'solid', color: '00FF00', pt: 0 }],
      },
    );
    generatedDifferenceSlide.addTable(
      [[{ text: 'PptxGenJS omitted type', options: {} }]],
      { x: 0.5, y: 3.5, w: 2, h: 0.5, border: { color: '112233', pt: 2 } },
    );
    const importedDifferences = await importPptxGenJS(generatedDifferences);
    const importedDifferenceTables = importedDifferences.slides[0]!.shapes.filter(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(borderMatrices(importedDifferenceTables)).toEqual([
      [four(line('666666', 1))],
      [four(line('666666', 1))],
      [{
        top: line('00FF00', 1),
        right: none,
        bottom: none,
        left: none,
      }],
      [four(line('112233', 2))],
    ]);

    const nativeDifferences = PptxDocument.create();
    const nativeDifferenceSlide = nativeDifferences.addSlide();
    const nativeEmptyTable = nativeDifferenceSlide.addTable(
      [['Native empty table']],
      { border: {} },
    );
    const nativeEmptyCell = nativeDifferenceSlide.addTable(
      [[{ text: 'Native empty cell', options: { border: {} } }]],
      {
        border: {
          kind: 'line',
          color: { kind: 'srgb', value: 'FF0000' },
          width: 2,
          style: 'solid',
        },
      },
    );
    const nativeTupleZero = nativeDifferenceSlide.addTable(
      [['Native tuple zero']],
      { border: [
        {
          kind: 'line',
          color: { kind: 'srgb', value: '00FF00' },
          width: 0,
          style: 'solid',
        },
        undefined,
        undefined,
        undefined,
      ] },
    );
    const nativeOmittedStyle = nativeDifferenceSlide.addTable(
      [['Native omitted style']],
      { border: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
        width: 2,
      } },
    );
    expect(nativeEmptyTable.rows[0]!.cells[0]!.borders).toEqual(four(none));
    expect(nativeEmptyCell.rows[0]!.cells[0]!.borders).toEqual(
      four(line('FF0000', 2)),
    );
    expect(nativeTupleZero.rows[0]!.cells[0]!.borders).toEqual({
      top: line('00FF00', 0),
      right: none,
      bottom: none,
      left: none,
    });
    const omittedStyleLine = {
      kind: 'line' as const,
      color: { kind: 'srgb' as const, value: '112233' },
      width: 2,
    };
    expect(nativeOmittedStyle.rows[0]!.cells[0]!.borders).toEqual(
      four(omittedStyleLine),
    );
    const nativeDifferenceXml = new TextDecoder().decode(
      nativeDifferences.opcPackage.requirePart(nativeDifferenceSlide.partUri).bytes,
    );
    const nativeDifferenceCells = [...nativeDifferenceXml.matchAll(
      /<a:tc>([\s\S]*?)<\/a:tc>/g,
    )].map((match) => match[1]!);
    expect(nativeDifferenceCells[0]!.match(/<a:noFill\/>/g)).toHaveLength(4);
    expect(nativeDifferenceCells[1]).toContain('<a:srgbClr val="FF0000"/>');
    expect(nativeDifferenceCells[2]).toContain(
      '<a:lnT w="0" cap="flat" cmpd="sng" algn="ctr">'
      + '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>',
    );
    expect(nativeDifferenceCells[3]).not.toContain('<a:prstDash');

    const beforeInvalid = native.opcPackage.requirePart(nativeSlide.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    expect(() => nativeSlide.addTable(
      [['Invalid short tuple']],
      { border: [{ kind: 'none' }] as never },
    )).toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeSlide.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
    expect(nativeSlide.shapes).toHaveLength(4);
  });

  it('locks bullet and numbering behavior across every declared owner', async () => {
    const numberTypes = [
      'alphaLcParenBoth',
      'alphaLcParenR',
      'alphaLcPeriod',
      'alphaUcParenBoth',
      'alphaUcParenR',
      'alphaUcPeriod',
      'arabicParenBoth',
      'arabicParenR',
      'arabicPeriod',
      'arabicPlain',
      'romanLcParenBoth',
      'romanLcParenR',
      'romanLcPeriod',
      'romanUcParenBoth',
      'romanUcParenR',
      'romanUcPeriod',
    ] as const;
    const bulletCases = [
      { name: 'boolean', value: true },
      { name: 'empty', value: {} },
      { name: 'typed-bullet', value: { type: 'bullet' } },
      { name: 'character', value: { characterCode: '25BA', indent: 18 } },
      { name: 'code-alias', value: { code: '25C6', indent: 19 } },
      {
        name: 'canonical-number',
        value: {
          type: 'number', numberType: 'romanUcPeriod', numberStartAt: 3, indent: 22,
        },
      },
      {
        name: 'legacy-number',
        value: { type: 'number', style: 'romanLcParenR', startAt: 4, indent: 24 },
      },
      { name: 'margin-alias', value: { marginPt: 4 } },
    ] as const;
    const source = new PptxGenJS();
    const slideNumberSlide = source.addSlide();
    slideNumberSlide.slideNumber = {
      x: 0,
      y: 0,
      w: 1,
      h: 0.3,
      bullet: {
        type: 'number', characterCode: '25BA', code: '25C6', indent: 18,
        marginPt: 4, numberType: 'romanUcPeriod', numberStartAt: 3,
        style: 'romanLcParenR', startAt: 4,
      },
    } as unknown as PptxGenJSSlideNumberProps;
    const slide = source.addSlide();
    for (const [index, fixture] of bulletCases.entries()) {
      slide.addText(fixture.name, {
        objectName: `bullet-text-${fixture.name}`,
        x: 0,
        y: index * 0.2,
        w: 2,
        h: 0.2,
        bullet: fixture.value,
      });
    }
    for (const [index, numberType] of numberTypes.entries()) {
      slide.addText(numberType, {
        objectName: `number-type-${numberType}`,
        x: 2,
        y: index * 0.2,
        w: 2,
        h: 0.2,
        bullet: { type: 'number', numberType, numberStartAt: index + 1 },
      });
      slide.addText(numberType, {
        objectName: `legacy-style-${numberType}`,
        x: 4,
        y: index * 0.2,
        w: 2,
        h: 0.2,
        bullet: { type: 'number', style: numberType, numberStartAt: index + 1 },
      });
    }
    slide.addTable([bulletCases.map((fixture) => ({
      text: fixture.name,
      options: { bullet: fixture.value },
    }))], {
      x: 0,
      y: 4,
      w: 8,
      h: 0.5,
      objectName: 'bullet-cell-matrix',
    });
    slide.addTable([['table boolean inert']], {
      x: 0,
      y: 5,
      w: 3,
      h: 0.5,
      objectName: 'table-bullet-boolean-inert',
      bullet: true,
    });
    slide.addTable([['table object inert']], {
      x: 3,
      y: 5,
      w: 3,
      h: 0.5,
      objectName: 'table-bullet-object-inert',
      bullet: {
        type: 'number', characterCode: '25BA', code: '25C6', indent: 18,
        marginPt: 4, numberType: 'romanUcPeriod', numberStartAt: 3,
        style: 'romanLcParenR', startAt: 4,
      },
    });
    source.defineSlideMaster({
      title: 'BULLET-OWNER-MATRIX',
      objects: bulletCases.map((fixture, index): PptxGenJSMasterObject => ({
        placeholder: {
          text: fixture.name,
          options: {
            name: `bullet-placeholder-${fixture.name}`,
            objectName: `bullet-placeholder-${fixture.name}`,
            type: 'body',
            x: 0,
            y: index * 0.3,
            w: 3,
            h: 0.3,
            bullet: fixture.value,
          },
        },
      })),
    });
    source.addSlide({ masterName: 'BULLET-OWNER-MATRIX' });

    const imported = await openPptxGenJSPublicOutput(source);
    const bulletSnapshot = (shape: ShapeModel) => shape.richText[0]?.bullet;
    const textShapes = new Map(imported.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .map((shape) => [shape.name, shape]));
    expect(bulletCases.map(({ name }) => bulletSnapshot(
      textShapes.get(`bullet-text-${name}`)!,
    ))).toEqual([
      { kind: 'bullet', character: '•', indent: 27 },
      { kind: 'bullet', character: '•', indent: 27 },
      undefined,
      { kind: 'bullet', character: '►', indent: 18 },
      { kind: 'bullet', character: '◆', indent: 19 },
      { kind: 'number', style: 'arabicPeriod', startAt: 3, indent: 22 },
      { kind: 'number', style: 'romanLcParenR', startAt: 4, indent: 24 },
      { kind: 'bullet', character: '•', indent: 27 },
    ]);
    expect(numberTypes.map((numberType) => bulletSnapshot(
      textShapes.get(`number-type-${numberType}`)!,
    ))).toEqual(numberTypes.map((_, index) => ({
      kind: 'number', style: 'arabicPeriod', startAt: index + 1, indent: 27,
    })));
    expect(numberTypes.map((numberType) => bulletSnapshot(
      textShapes.get(`legacy-style-${numberType}`)!,
    ))).toEqual(numberTypes.map((style, index) => ({
      kind: 'number', style, startAt: index + 1, indent: 27,
    })));

    const matrixXml = slideXml(imported, 1);
    const tableFrames = [...matrixXml.matchAll(
      /<p:graphicFrame(?:\s[^>]*)?>[\s\S]*?<\/p:graphicFrame>/gu,
    )].map((match) => match[0]);
    expect(tableFrames).toHaveLength(3);
    const activeCells = [...tableFrames[0]!.matchAll(
      /<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/gu,
    )].map((match) => match[0]);
    expect(activeCells).toHaveLength(8);
    expect(activeCells[0]).toContain('<a:buChar char="&#x2022;"/>');
    expect(activeCells[1]).toContain('<a:buChar char="&#x2022;"/>');
    expect(activeCells[2]).not.toMatch(/<a:bu(?:Char|AutoNum)\b/u);
    expect(activeCells[3]).toContain('<a:buChar char="&#x25BA;"/>');
    expect(activeCells[3]).toContain('marL="228600" indent="-228600"');
    expect(activeCells[4]).toContain('<a:buChar char="&#x25C6;"/>');
    expect(activeCells[5]).toContain('<a:buAutoNum type="arabicPeriod" startAt="3"/>');
    expect(activeCells[6]).toContain('<a:buAutoNum type="romanLcParenR" startAt="4"/>');
    expect(activeCells[7]).toContain('<a:buChar char="&#x2022;"/>');
    expect(tableFrames.slice(1).every(
      (frame) => !/<a:bu(?:Char|AutoNum)\b/u.test(frame),
    )).toBe(true);

    const layout = imported.layouts.find(({ name }) => name === 'BULLET-OWNER-MATRIX')!;
    const placeholders = new Map(layout.placeholders
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .map((shape) => [shape.name, shape]));
    expect(bulletCases.map(({ name }) => bulletSnapshot(
      placeholders.get(`bullet-placeholder-${name}`)!,
    ))).toEqual([
      { kind: 'bullet', character: '•', indent: 27 },
      { kind: 'bullet', character: '•', indent: 27 },
      undefined,
      { kind: 'bullet', character: '►', indent: 18 },
      { kind: 'bullet', character: '◆', indent: 19 },
      { kind: 'number', style: 'arabicPeriod', startAt: 3, indent: 22 },
      { kind: 'number', style: 'romanLcParenR', startAt: 4, indent: 24 },
      { kind: 'bullet', character: '•', indent: 27 },
    ]);
    const slideNumberXml = slideNumberOwnerState(imported, imported.slides[0]!.partUri).xml
      .match(/<p:sp(?:\s[^>]*)?>[\s\S]*?<p:ph\b[^>]*\btype="sldNum"[\s\S]*?<\/p:sp>/u)?.[0];
    expect(slideNumberXml).toBeDefined();
    expect(slideNumberXml).not.toMatch(/<a:bu(?:Char|AutoNum)\b/u);

    const generateTableToSlides = async (bullet?: unknown) => {
      const presentation = new PptxGenJS();
      await withTableToSlidesGlobals(tableToSlidesFixture(), () =>
        presentation.tableToSlides('report', bullet === undefined ? {} : { bullet }));
      return openPptxGenJSPublicOutput(presentation);
    };
    const tableToSlidesBaseline = await generateTableToSlides();
    const tableToSlidesVariant = await generateTableToSlides({
      type: 'number', characterCode: '25BA', code: '25C6', indent: 18,
      marginPt: 4, numberType: 'romanUcPeriod', numberStartAt: 3,
      style: 'romanLcParenR', startAt: 4,
    });
    expect(slideXml(tableToSlidesVariant, 0)).toBe(slideXml(tableToSlidesBaseline, 0));
  });

  it('locks underline behavior across every declared owner', async () => {
    const underlineStyles = [
      'dash',
      'dashHeavy',
      'dashLong',
      'dashLongHeavy',
      'dbl',
      'dotDash',
      'dotDashHeave',
      'dotDotDash',
      'dotDotDashHeavy',
      'dotted',
      'dottedHeavy',
      'heavy',
      'none',
      'sng',
      'wavy',
      'wavyDbl',
      'wavyHeavy',
    ] as const;
    const underlineCases = [
      ...underlineStyles.map((style) => ({ name: style, value: { style } })),
      { name: 'styled-color', value: { style: 'dbl', color: 'FF0000' } },
      { name: 'color-only', value: { color: '00FF00' } },
      { name: 'empty-object', value: {} },
      { name: 'boolean-true', value: true },
      { name: 'boolean-false', value: false },
    ] as const;
    const namedOwnerXml = (
      xml: string,
      name: string,
      ownerTag: 'p:sp' | 'p:graphicFrame',
    ) => {
      const nameOffset = xml.indexOf(`name="${name}"`);
      expect(nameOffset).toBeGreaterThanOrEqual(0);
      const start = xml.lastIndexOf(`<${ownerTag}`, nameOffset);
      const end = xml.indexOf(`</${ownerTag}>`, nameOffset);
      expect(start).toBeGreaterThanOrEqual(0);
      expect(end).toBeGreaterThan(nameOffset);
      return xml.slice(start, end + ownerTag.length + 3);
    };
    const underlineTokens = (xml: string) => [...xml.matchAll(
      /<a:(?:rPr|defRPr)\b[^>]*\bu="([^"]+)"/gu,
    )].map((match) => match[1]);

    const source = new PptxGenJS();
    const activeSlide = source.addSlide();
    activeSlide.addText(
      underlineStyles.map((style) => ({ text: style, options: { underline: { style } } })),
      {
        objectName: 'underline-text-owner',
        x: 0,
        y: 0,
        w: 6,
        h: 1,
      },
    );
    activeSlide.addTable([underlineCases.map(({ name, value }) => ({
      text: name,
      options: { underline: value },
    }))], {
      objectName: 'underline-cell-owner',
      x: 0,
      y: 1,
      w: 12,
      h: 0.5,
    });
    underlineCases.forEach(({ name, value }, index) => {
      activeSlide.addTable([[name]], {
        objectName: `underline-table-owner-${name}`,
        x: index % 6,
        y: 2 + Math.floor(index / 6) * 0.4,
        w: 1,
        h: 0.4,
        underline: value,
      });
    });
    activeSlide.addText([
      { text: 'styled-color', options: { underline: { style: 'dbl', color: 'FF0000' } } },
      { text: 'color-only', options: { underline: { color: '00FF00' } } },
      { text: 'empty', options: { underline: {} } },
      {
        text: 'runtime-correct',
        options: { underline: { style: 'dotDashHeavy' } as unknown },
      },
      { text: 'boolean-true', options: { underline: true } },
      { text: 'boolean-false', options: { underline: false } },
    ], {
      objectName: 'underline-edge-cases',
      x: 0,
      y: 4,
      w: 8,
      h: 0.5,
    });
    source.defineSlideMaster({
      title: 'UNDERLINE-OWNER-MATRIX',
      objects: underlineCases.map(({ name, value }, index): PptxGenJSMasterObject => ({
        placeholder: {
          text: name,
          options: {
            name: `underline-placeholder-${name}`,
            objectName: `underline-placeholder-${name}`,
            type: 'body',
            x: index % 6,
            y: Math.floor(index / 6) * 0.4,
            w: 1,
            h: 0.4,
            underline: value,
          },
        },
      })),
    });
    source.addSlide({ masterName: 'UNDERLINE-OWNER-MATRIX' });
    for (const [index, { value: underline }] of underlineCases.entries()) {
      const slide = source.addSlide();
      slide.addText(`Slide number control ${index + 1}`, {
        objectName: `underline-slide-number-control-${index}`,
        x: 0,
        y: 0,
        w: 3,
        h: 0.3,
        bold: true,
      });
      slide.slideNumber = {
        x: 0,
        y: 0.4,
        w: 1,
        h: 0.3,
        underline,
      } as unknown as PptxGenJSSlideNumberProps;
    }

    const imported = await openPptxGenJSPublicOutput(source);
    const activeXml = slideXml(imported, 0);
    expect(underlineTokens(namedOwnerXml(
      activeXml,
      'underline-text-owner',
      'p:sp',
    ))).toEqual(underlineStyles);
    const cellOwnerXml = namedOwnerXml(
      activeXml,
      'underline-cell-owner',
      'p:graphicFrame',
    );
    expect(underlineTokens(cellOwnerXml)).toEqual([...underlineStyles, 'dbl']);
    expect(cellOwnerXml).toContain(
      '<a:uFill><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:uFill>',
    );
    expect(cellOwnerXml).toContain(
      '<a:uFill><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:uFill>',
    );
    expect(underlineCases.map(({ name }) => underlineTokens(namedOwnerXml(
      activeXml,
      `underline-table-owner-${name}`,
      'p:graphicFrame',
    ))[0])).toEqual([
      ...underlineStyles,
      'dbl',
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    const tableStyledColorXml = namedOwnerXml(
      activeXml,
      'underline-table-owner-styled-color',
      'p:graphicFrame',
    );
    const tableColorOnlyXml = namedOwnerXml(
      activeXml,
      'underline-table-owner-color-only',
      'p:graphicFrame',
    );
    expect(tableStyledColorXml).toContain(
      '<a:uFill><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:uFill>',
    );
    expect(tableColorOnlyXml).not.toMatch(/<a:rPr\b[^>]*\bu=/u);
    expect(tableColorOnlyXml).toContain(
      '<a:uFill><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:uFill>',
    );
    expect(namedOwnerXml(
      activeXml,
      'underline-table-owner-empty-object',
      'p:graphicFrame',
    )).not.toMatch(/\bu=|<a:uFill>/u);

    const layout = imported.layouts.find(({ name }) => name === 'UNDERLINE-OWNER-MATRIX')!;
    const layoutXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(layout.partUri).bytes,
    );
    expect(underlineCases.map(({ name }) => {
      const tokens = underlineTokens(namedOwnerXml(
        layoutXml,
        `underline-placeholder-${name}`,
        'p:sp',
      ));
      const expected = name === 'styled-color'
        ? ['dbl', 'dbl']
        : name === 'boolean-true'
        ? ['sng', 'sng']
          : ['boolean-false', 'color-only', 'empty-object'].includes(name)
            ? [] : [name, name];
      expect(tokens).toEqual(expected);
      return tokens[0];
    })).toEqual([
      ...underlineStyles,
      'dbl',
      undefined,
      undefined,
      'sng',
      undefined,
    ]);
    const placeholderStyledColorXml = namedOwnerXml(
      layoutXml,
      'underline-placeholder-styled-color',
      'p:sp',
    );
    const placeholderColorOnlyXml = namedOwnerXml(
      layoutXml,
      'underline-placeholder-color-only',
      'p:sp',
    );
    expect(placeholderStyledColorXml.match(/<a:uFill>/gu)).toHaveLength(2);
    expect(placeholderStyledColorXml).toContain('<a:srgbClr val="FF0000"/>');
    expect(placeholderColorOnlyXml).not.toMatch(/<a:(?:rPr|defRPr)\b[^>]*\bu=/u);
    expect(placeholderColorOnlyXml.match(/<a:uFill>/gu)).toHaveLength(2);
    expect(placeholderColorOnlyXml).toContain('<a:srgbClr val="00FF00"/>');

    const edgeRuns = [...namedOwnerXml(
      activeXml,
      'underline-edge-cases',
      'p:sp',
    ).matchAll(/<a:r>([\s\S]*?)<\/a:r>/gu)].map((match) => match[1]!);
    expect(edgeRuns).toHaveLength(6);
    expect(edgeRuns[0]).toContain('u="dbl"');
    expect(edgeRuns[0]).toContain(
      '<a:uFill><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:uFill>',
    );
    expect(edgeRuns[1]).not.toMatch(/<a:rPr\b[^>]*\bu=/u);
    expect(edgeRuns[1]).toContain(
      '<a:uFill><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:uFill>',
    );
    expect(edgeRuns[2]).not.toMatch(/\bu=|<a:uFill>/u);
    expect(edgeRuns[3]).toContain('u="dotDashHeavy"');
    expect(edgeRuns[4]).toContain('u="sng"');
    expect(edgeRuns[5]).not.toMatch(/<a:rPr\b[^>]*\bu=/u);

    for (let index = 0; index < underlineCases.length; index += 1) {
      const slideXmlValue = slideXml(imported, index + 2);
      const slideNumberXml = slideNumberOwnerState(
        imported,
        imported.slides[index + 2]!.partUri,
      ).xml.match(
        /<p:sp(?:\s[^>]*)?>[\s\S]*?<p:ph\b[^>]*\btype="sldNum"[\s\S]*?<\/p:sp>/u,
      )?.[0];
      expect(slideXmlValue).toContain(`underline-slide-number-control-${index}`);
      expect(slideNumberXml).toBeDefined();
      expect(slideNumberXml).not.toMatch(/<a:(?:rPr|defRPr)\b[^>]*\bu=|<a:uFill>/u);
    }

    const generateTableToSlides = async (underline?: unknown) => {
      const presentation = new PptxGenJS();
      await withTableToSlidesGlobals(tableToSlidesFixture(), () =>
        presentation.tableToSlides(
          'report',
          underline === undefined ? {} : { underline },
        ));
      return openPptxGenJSPublicOutput(presentation);
    };
    const tableToSlidesBaseline = await generateTableToSlides();
    const tableToSlidesBaselineXml = slideXml(tableToSlidesBaseline, 0);
    for (const { value: underline } of underlineCases) {
      const variant = await generateTableToSlides(underline);
      expect(slideXml(variant, 0)).toBe(tableToSlidesBaselineXml);
    }
  });

  it('locks scalar text formatting behavior across every declared owner', async () => {
    const style = {
      bold: true,
      color: '112233',
      fontFace: 'ProbeFace',
      fontSize: 24,
      highlight: 'FFFF00',
      italic: true,
      lang: 'fr-CA',
    } as const;
    const namedOwnerXml = (
      xml: string,
      name: string,
      ownerTag: 'p:sp' | 'p:graphicFrame',
    ) => {
      const nameOffset = xml.indexOf(`name="${name}"`);
      expect(nameOffset).toBeGreaterThanOrEqual(0);
      const start = xml.lastIndexOf(`<${ownerTag}`, nameOffset);
      const end = xml.indexOf(`</${ownerTag}>`, nameOffset);
      expect(start).toBeGreaterThanOrEqual(0);
      expect(end).toBeGreaterThan(nameOffset);
      return xml.slice(start, end + ownerTag.length + 3);
    };
    const paragraphCount = (xml: string) => xml.match(/<a:p>/gu)?.length ?? 0;
    const assertScalarStyle = (
      xml: string,
      expected: {
        readonly color: string;
        readonly fontFace: string;
        readonly fontSize: number;
        readonly highlight?: string;
        readonly italic?: boolean;
        readonly lang: string;
      },
    ) => {
      expect(xml).toContain(`lang="${expected.lang}"`);
      expect(xml).toContain(`sz="${expected.fontSize * 100}"`);
      expect(xml).toContain('b="1"');
      if (expected.italic) expect(xml).toContain('i="1"');
      expect(xml).toContain(`val="${expected.color}"`);
      expect(xml).toContain(`typeface="${expected.fontFace}"`);
      if (expected.highlight) {
        expect(xml).toContain(`<a:highlight><a:srgbClr val="${expected.highlight}"/></a:highlight>`);
      }
    };

    const source = new PptxGenJS();
    const slide = source.addSlide();
    slide.addText([
      { text: 'Text first', options: { ...style, breakLine: true } },
      { text: 'Text second', options: {} },
      { text: 'Text soft', options: { softBreakBefore: true } },
    ], {
      objectName: 'scalar-text-owner',
      x: 0,
      y: 0,
      w: 4,
      h: 1,
    });
    slide.addTable([['Table defaults']], {
      objectName: 'scalar-table-owner',
      x: 0,
      y: 1.25,
      w: 4,
      h: 0.5,
      ...style,
      breakLine: true,
      softBreakBefore: true,
    });
    slide.addTable([[
      {
        text: 'Cell styled',
        options: {
          ...style,
          bold: false,
          color: '223344',
          fontFace: 'CellFace',
          fontSize: 20,
          highlight: '00FF00',
          lang: 'de-DE',
        },
      },
      {
        text: [{ text: 'Break one' }, { text: 'Break two' }],
        options: { breakLine: true },
      },
      {
        text: [{ text: 'Soft one' }, { text: 'Soft two' }],
        options: { softBreakBefore: true },
      },
    ]], {
      objectName: 'scalar-cell-owner',
      x: 0,
      y: 2,
      w: 9,
      h: 0.5,
      bold: true,
    });
    source.defineSlideMaster({
      title: 'SCALAR-TEXT-OWNER-MATRIX',
      objects: [{
        placeholder: {
          text: 'Placeholder scalar style',
          options: {
            name: 'scalar-placeholder-owner',
            objectName: 'scalar-placeholder-owner',
            type: 'body',
            x: 0,
            y: 0,
            w: 4,
            h: 1,
            ...style,
            breakLine: true,
            softBreakBefore: true,
          },
        },
      }],
    });
    source.addSlide({ masterName: 'SCALAR-TEXT-OWNER-MATRIX' });
    const slideNumberSlide = source.addSlide();
    slideNumberSlide.slideNumber = {
      x: 0,
      y: 0,
      w: 1,
      h: 0.5,
      ...style,
      breakLine: true,
      softBreakBefore: true,
    } as unknown as PptxGenJSSlideNumberProps;

    const imported = await openPptxGenJSPublicOutput(source);
    const activeXml = slideXml(imported, 0);
    const textXml = namedOwnerXml(activeXml, 'scalar-text-owner', 'p:sp');
    assertScalarStyle(textXml, style);
    expect(paragraphCount(textXml)).toBe(2);
    expect(textXml.match(/<a:br\/>/gu)).toHaveLength(1);

    const tableXml = namedOwnerXml(activeXml, 'scalar-table-owner', 'p:graphicFrame');
    assertScalarStyle(tableXml, {
      color: '112233',
      fontFace: 'ProbeFace',
      fontSize: 24,
      lang: 'en-US',
    });
    expect(tableXml).not.toContain('i="1"');
    expect(tableXml).not.toContain('<a:highlight>');
    expect(tableXml).not.toContain('lang="fr-CA"');
    expect(paragraphCount(tableXml)).toBe(1);
    expect(tableXml).not.toContain('<a:br/>');

    const cellTableXml = namedOwnerXml(activeXml, 'scalar-cell-owner', 'p:graphicFrame');
    const cellXml = [...cellTableXml.matchAll(/<a:tc(?:\s[^>]*)?>([\s\S]*?)<\/a:tc>/gu)]
      .map((match) => match[1]!);
    expect(cellXml).toHaveLength(3);
    assertScalarStyle(cellXml[0]!, {
      color: '223344',
      fontFace: 'CellFace',
      fontSize: 20,
      highlight: '00FF00',
      italic: true,
      lang: 'de-DE',
    });
    expect(cellXml[0]).not.toContain('b="0"');
    expect(paragraphCount(cellXml[1]!)).toBe(2);
    expect(paragraphCount(cellXml[2]!)).toBe(1);
    expect(cellXml[2]!.match(/<a:br\/>/gu)).toHaveLength(1);

    const layout = imported.layouts.find(
      ({ name }) => name === 'SCALAR-TEXT-OWNER-MATRIX',
    )!;
    const layoutXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(layout.partUri).bytes,
    );
    const placeholderXml = namedOwnerXml(
      layoutXml,
      'scalar-placeholder-owner',
      'p:sp',
    );
    assertScalarStyle(placeholderXml, style);
    expect(paragraphCount(placeholderXml)).toBe(1);
    expect(placeholderXml).not.toContain('<a:br/>');

    const slideNumberXml = slideNumberOwnerState(
      imported,
      imported.slides[2]!.partUri,
    ).xml;
    expect(slideNumberXml).toContain('<a:defRPr sz="2400">');
    expect(slideNumberXml).toContain('<a:srgbClr val="112233"/>');
    expect(slideNumberXml).toContain('<a:latin typeface="ProbeFace"/>');
    expect(slideNumberXml).toContain('<a:rPr b="1" lang="en-US"/>');
    expect(slideNumberXml).not.toContain('i="1"');
    expect(slideNumberXml).not.toContain('<a:highlight>');
    expect(slideNumberXml).not.toContain('lang="fr-CA"');
    expect(slideNumberXml).not.toContain('<a:br/>');

    const tableToSlidesVariants = [
      ['bold', true],
      ['breakLine', true],
      ['color', 'A1B2C3'],
      ['fontFace', 'IgnoredFace'],
      ['fontSize', 30],
      ['highlight', 'ABCDEF'],
      ['italic', true],
      ['lang', 'ja-JP'],
      ['softBreakBefore', true],
    ] as const;
    const generateTableToSlides = async (options: Record<string, unknown>) => {
      const presentation = new PptxGenJS();
      await withTableToSlidesGlobals(tableToSlidesFixture(), () =>
        presentation.tableToSlides('report', options));
      return openPptxGenJSPublicOutput(presentation);
    };
    const baseline = await generateTableToSlides({});
    const baselineXml = slideXml(baseline, 0);
    expect(baselineXml).toContain('val="010203"');
    expect(baselineXml).toContain('typeface="Arial"');
    expect(baselineXml).toContain('sz="1200"');
    expect(baselineXml).toContain('b="1"');
    for (const [property, value] of tableToSlidesVariants) {
      const variant = await generateTableToSlides({ [property]: value });
      expect(slideXml(variant, 0), property).toBe(baselineXml);
    }
  }, 20_000);

  it('locks text direction behavior across every declared owner', async () => {
    const tableDirections = ['horz', 'vert', 'vert270', 'wordArtVert'] as const;
    const textDirections = [
      'eaVert',
      'horz',
      'mongolianVert',
      'vert',
      'vert270',
      'wordArtVert',
      'wordArtVertRtl',
    ] as const;
    const namedOwnerXml = (
      xml: string,
      name: string,
      ownerTag: 'p:sp' | 'p:graphicFrame',
    ) => {
      const nameOffset = xml.indexOf(`name="${name}"`);
      expect(nameOffset).toBeGreaterThanOrEqual(0);
      const start = xml.lastIndexOf(`<${ownerTag}`, nameOffset);
      const end = xml.indexOf(`</${ownerTag}>`, nameOffset);
      expect(start).toBeGreaterThanOrEqual(0);
      expect(end).toBeGreaterThan(nameOffset);
      return xml.slice(start, end + ownerTag.length + 3);
    };
    const bodyDirection = (xml: string) =>
      xml.match(/<a:bodyPr\b[^>]*\bvert="([^"]+)"/u)?.[1];
    const cellDirections = (xml: string) => [...xml.matchAll(
      /<a:tc(?:\s[^>]*)?>([\s\S]*?)<\/a:tc>/gu,
    )].map((match) =>
      match[1]!.match(/<a:tcPr\b[^>]*\bvert="([^"]+)"/u)?.[1]);

    const source = new PptxGenJS();
    const activeSlide = source.addSlide();
    activeSlide.addTable([tableDirections.map((textDirection) => ({
      text: `Cell ${textDirection}`,
      options: { textDirection },
    }))], {
      objectName: 'text-direction-cell-owner',
      x: 0,
      y: 0,
      w: 12,
      h: 0.5,
    });
    tableDirections.forEach((textDirection, index) => {
      activeSlide.addTable([[`Table ${textDirection}`]], {
        objectName: `text-direction-table-owner-${textDirection}`,
        x: index * 3,
        y: 0.75,
        w: 3,
        h: 0.5,
        textDirection,
      });
    });
    textDirections.forEach((vert, index) => {
      activeSlide.addText(`Text ${vert}`, {
        objectName: `text-direction-vert-owner-${vert}`,
        x: index % 4 * 3,
        y: 1.5 + Math.floor(index / 4) * 0.5,
        w: 3,
        h: 0.5,
        vert,
      });
    });
    tableDirections.forEach((textDirection, index) => {
      activeSlide.addText(`Ignored ${textDirection}`, {
        objectName: `text-direction-text-owner-${textDirection}`,
        x: index * 3,
        y: 2.75,
        w: 3,
        h: 0.5,
        textDirection,
      });
    });
    activeSlide.addText('Canonical direction wins', {
      objectName: 'text-direction-canonical-wins',
      x: 0,
      y: 3.5,
      w: 4,
      h: 0.5,
      textDirection: 'vert',
      vert: 'vert270',
    });
    activeSlide.addText(
      [{
        text: 'Run direction is ignored',
        options: { vert: 'wordArtVert', textDirection: 'vert' } as Record<string, unknown>,
      }],
      {
        objectName: 'text-direction-run-owner',
        x: 4,
        y: 3.5,
        w: 4,
        h: 0.5,
      },
    );

    source.defineSlideMaster({
      title: 'TEXT-DIRECTION-OWNER-MATRIX',
      objects: tableDirections.map((textDirection, index): PptxGenJSMasterObject => ({
        placeholder: {
          text: `Placeholder ${textDirection}`,
          options: {
            name: `text-direction-placeholder-owner-${textDirection}`,
            objectName: `text-direction-placeholder-owner-${textDirection}`,
            type: 'body',
            x: index * 3,
            y: 0,
            w: 3,
            h: 0.5,
            textDirection,
          },
        },
      })),
    });
    source.addSlide({ masterName: 'TEXT-DIRECTION-OWNER-MATRIX' });
    for (const [index, textDirection] of tableDirections.entries()) {
      const slide = source.addSlide();
      slide.addText(`Slide number control ${textDirection}`, {
        objectName: `text-direction-slide-number-control-${textDirection}`,
        x: 0,
        y: 0,
        w: 4,
        h: 0.5,
      });
      slide.slideNumber = {
        x: 0,
        y: 0.5,
        w: 1,
        h: 0.5,
        textDirection,
        bold: index % 2 === 0,
      } as PptxGenJSSlideNumberProps;
    }

    const imported = await openPptxGenJSPublicOutput(source);
    const activeXml = slideXml(imported, 0);
    expect(cellDirections(namedOwnerXml(
      activeXml,
      'text-direction-cell-owner',
      'p:graphicFrame',
    ))).toEqual([undefined, 'vert', 'vert270', 'wordArtVert']);
    expect(tableDirections.map((textDirection) => cellDirections(namedOwnerXml(
      activeXml,
      `text-direction-table-owner-${textDirection}`,
      'p:graphicFrame',
    ))[0])).toEqual([undefined, 'vert', 'vert270', 'wordArtVert']);
    expect(textDirections.map((vert) => bodyDirection(namedOwnerXml(
      activeXml,
      `text-direction-vert-owner-${vert}`,
      'p:sp',
    )))).toEqual(textDirections);
    for (const textDirection of tableDirections) {
      expect(bodyDirection(namedOwnerXml(
        activeXml,
        `text-direction-text-owner-${textDirection}`,
        'p:sp',
      ))).toBeUndefined();
    }
    expect(bodyDirection(namedOwnerXml(
      activeXml,
      'text-direction-canonical-wins',
      'p:sp',
    ))).toBe('vert270');
    expect(bodyDirection(namedOwnerXml(
      activeXml,
      'text-direction-run-owner',
      'p:sp',
    ))).toBeUndefined();

    const layout = imported.layouts.find(
      ({ name }) => name === 'TEXT-DIRECTION-OWNER-MATRIX',
    )!;
    const layoutXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(layout.partUri).bytes,
    );
    for (const textDirection of tableDirections) {
      const placeholderXml = namedOwnerXml(
        layoutXml,
        `text-direction-placeholder-owner-${textDirection}`,
        'p:sp',
      );
      expect(placeholderXml).toContain(`Placeholder ${textDirection}`);
      expect(bodyDirection(placeholderXml)).toBeUndefined();
    }

    for (const [index, textDirection] of tableDirections.entries()) {
      const slideXmlValue = slideXml(imported, index + 2);
      const slideNumberXml = slideNumberOwnerState(
        imported,
        imported.slides[index + 2]!.partUri,
      ).xml.match(
        /<p:sp(?:\s[^>]*)?>[\s\S]*?<p:ph\b[^>]*\btype="sldNum"[\s\S]*?<\/p:sp>/u,
      )?.[0];
      expect(slideXmlValue).toContain(
        `text-direction-slide-number-control-${textDirection}`,
      );
      expect(slideNumberXml).toBeDefined();
      expect(bodyDirection(slideNumberXml!)).toBeUndefined();
    }

    const generateTableToSlides = async (textDirection?: unknown) => {
      const presentation = new PptxGenJS();
      await withTableToSlidesGlobals(tableToSlidesFixture(), () =>
        presentation.tableToSlides(
          'report',
          textDirection === undefined ? {} : { textDirection },
        ));
      return openPptxGenJSPublicOutput(presentation);
    };
    const tableToSlidesBaseline = await generateTableToSlides();
    const tableToSlidesBaselineXml = slideXml(tableToSlidesBaseline, 0);
    expect(tableToSlidesBaselineXml).toContain('<a:tbl>');
    for (const textDirection of tableDirections) {
      const variant = await generateTableToSlides(textDirection);
      expect(slideXml(variant, 0)).toBe(tableToSlidesBaselineXml);
    }
  });

  it('locks tab stop behavior across every declared owner', async () => {
    const tabStops = [
      { position: 1, alignment: 'l' },
      { position: 1.25, alignment: 'r' },
      { position: 1.5, alignment: 'ctr' },
      { position: 1.75, alignment: 'dec' },
    ] as const;
    const expectedStops = [
      { position: 1, alignment: 'left' },
      { position: 1.25, alignment: 'right' },
      { position: 1.5, alignment: 'center' },
      { position: 1.75, alignment: 'decimal' },
    ];
    const expectedXml = '<a:tab pos="914400" algn="l"/>'
      + '<a:tab pos="1143000" algn="r"/>'
      + '<a:tab pos="1371600" algn="ctr"/>'
      + '<a:tab pos="1600200" algn="dec"/>';
    const namedOwnerXml = (xml: string, name: string, ownerTag: 'p:sp' | 'p:graphicFrame') => {
      const nameOffset = xml.indexOf(`name="${name}"`);
      expect(nameOffset).toBeGreaterThanOrEqual(0);
      const start = xml.lastIndexOf(`<${ownerTag}`, nameOffset);
      const end = xml.indexOf(`</${ownerTag}>`, nameOffset);
      expect(start).toBeGreaterThanOrEqual(0);
      expect(end).toBeGreaterThan(nameOffset);
      return xml.slice(start, end + ownerTag.length + 3);
    };

    const source = new PptxGenJS();
    const slideNumberSlide = source.addSlide();
    slideNumberSlide.slideNumber = {
      x: 0,
      y: 0,
      w: 1,
      h: 0.3,
      tabStops,
    } as unknown as PptxGenJSSlideNumberProps;
    const slide = source.addSlide();
    slide.addText('Text\t12.50', {
      objectName: 'tab-stops-text',
      x: 0,
      y: 0,
      w: 4,
      h: 0.5,
      tabStops,
    });
    slide.addTable([[{
      text: 'Cell\t42.75',
      options: { tabStops },
    }]], {
      objectName: 'tab-stops-cell',
      x: 0,
      y: 1,
      w: 4,
      h: 0.5,
    });
    slide.addTable([['Table default\tignored']], {
      objectName: 'tab-stops-table-inert',
      x: 0,
      y: 2,
      w: 4,
      h: 0.5,
      tabStops,
    });
    source.defineSlideMaster({
      title: 'TAB-STOPS-OWNER-MATRIX',
      objects: [{
        placeholder: {
          text: 'Placeholder\t100',
          options: {
            name: 'tab-stops-placeholder',
            objectName: 'tab-stops-placeholder',
            type: 'body',
            x: 0,
            y: 0,
            w: 4,
            h: 0.5,
            tabStops,
          },
        },
      }],
    });
    source.addSlide({ masterName: 'TAB-STOPS-OWNER-MATRIX' });

    const imported = await openPptxGenJSPublicOutput(source);
    const text = imported.slides[1]!.shapes.find(
      (shape): shape is ShapeModel => shape instanceof ShapeModel
        && shape.name === 'tab-stops-text',
    )!;
    const cellTable = imported.slides[1]!.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel
        && shape.name === 'tab-stops-cell',
    )!;
    const layout = imported.layouts.find(({ name }) => name === 'TAB-STOPS-OWNER-MATRIX')!;
    const placeholder = layout.placeholders.find(
      (shape): shape is ShapeModel => shape instanceof ShapeModel
        && shape.name === 'tab-stops-placeholder',
    )!;
    expect(text.richText[0]!.tabStops).toEqual(expectedStops);
    expect(cellTable.rows[0]!.cells[0]!.richText[0]!.tabStops).toEqual(expectedStops);
    expect(placeholder.text).toBe('Placeholder\t100');

    const ownerSlideXml = slideXml(imported, 1);
    const textXml = namedOwnerXml(ownerSlideXml, 'tab-stops-text', 'p:sp');
    const cellXml = namedOwnerXml(ownerSlideXml, 'tab-stops-cell', 'p:graphicFrame');
    const inertTableXml = namedOwnerXml(
      ownerSlideXml,
      'tab-stops-table-inert',
      'p:graphicFrame',
    );
    const layoutXml = new TextDecoder().decode(
      imported.opcPackage.requirePart(layout.partUri).bytes,
    );
    const placeholderXml = namedOwnerXml(layoutXml, 'tab-stops-placeholder', 'p:sp');
    expect(textXml).toContain(expectedXml);
    expect(cellXml).toContain(expectedXml);
    expect(placeholderXml).toContain(expectedXml);
    expect(inertTableXml).not.toContain('<a:tabLst>');

    const slideNumberXml = slideNumberOwnerState(imported, imported.slides[0]!.partUri).xml
      .match(/<p:sp(?:\s[^>]*)?>[\s\S]*?<p:ph\b[^>]*\btype="sldNum"[\s\S]*?<\/p:sp>/u)?.[0];
    expect(slideNumberXml).toBeDefined();
    expect(slideNumberXml).not.toContain('<a:tabLst>');

    const generateTableToSlides = async (stops?: unknown) => {
      const presentation = new PptxGenJS();
      await withTableToSlidesGlobals(tableToSlidesFixture(), () =>
        presentation.tableToSlides('report', stops === undefined ? {} : { tabStops: stops }));
      return openPptxGenJSPublicOutput(presentation);
    };
    const tableToSlidesBaseline = await generateTableToSlides();
    const tableToSlidesVariant = await generateTableToSlides(tabStops);
    expect(slideXml(tableToSlidesVariant, 0)).toBe(slideXml(tableToSlidesBaseline, 0));
  });

  it('imports public PptxGenJS output and continues editing in the OOXML kernel', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const generatedSlide = generated.addSlide();
    generatedSlide.addText('Created by PptxGenJS', { x: 1, y: 1, w: 7, h: 1, align: 'center' });
    generatedSlide.addText(
      [
        {
          text: 'Bold red',
          options: { bold: true, fontFace: 'Aptos', fontSize: 24, color: 'ff0000' },
        },
        {
          text: 'italic',
          options: { italic: true, fontSize: 14, color: '4472C4', softBreakBefore: true },
        },
      ],
      { x: 1, y: 2, w: 7, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Left', options: { align: 'left' } },
        { text: 'Center', options: { align: 'center' } },
        { text: 'Right', options: { align: 'right' } },
        { text: 'Justify', options: { align: 'justify' } },
      ],
      { x: 1, y: 3, w: 7, h: 2, align: 'left' },
    );
    generatedSlide.addText('Standard\nSecond', { x: 1, y: 5, w: 3, h: 1, bullet: true });
    generatedSlide.addText('Custom', {
      x: 4,
      y: 5,
      w: 3,
      h: 1,
      bullet: { characterCode: '25BA', indent: 18 },
    });
    generatedSlide.addText('Public numberType', {
      x: 7,
      y: 5,
      w: 3,
      h: 1,
      bullet: { type: 'number', numberType: 'romanUcPeriod', numberStartAt: 3, indent: 22 },
    });
    generatedSlide.addText('Deprecated style', {
      x: 10,
      y: 5,
      w: 2,
      h: 1,
      bullet: { type: 'number', style: 'romanLcParenR', startAt: 4, indent: 24 },
    });
    generatedSlide.addText('Exact first\nExact second', {
      x: 1,
      y: 6,
      w: 3,
      h: 1,
      lineSpacing: 28,
      lineSpacingMultiple: 1.5,
      paraSpaceBefore: 6.25,
      paraSpaceAfter: 8.5,
    });
    generatedSlide.addText('Multiple', {
      x: 4,
      y: 6,
      w: 3,
      h: 1,
      lineSpacingMultiple: 1.5,
      paraSpaceBefore: 4.25,
      paraSpaceAfter: 7.75,
    });
    generatedSlide.addText('Zero spacing', {
      x: 7,
      y: 6,
      w: 3,
      h: 1,
      lineSpacing: 0,
      lineSpacingMultiple: 0,
      paraSpaceBefore: 0,
      paraSpaceAfter: 0,
    });
    generatedSlide.addText('Level one', {
      x: 10,
      y: 6,
      w: 2,
      h: 0.5,
      bullet: true,
      indentLevel: 1,
    });
    generatedSlide.addText('Custom level two', {
      x: 10,
      y: 6.5,
      w: 2,
      h: 0.5,
      bullet: { characterCode: '25BA', indent: 18 },
      indentLevel: 2,
    });
    generatedSlide.addText('Number level three', {
      x: 10,
      y: 7,
      w: 2,
      h: 0.5,
      bullet: { type: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
      indentLevel: 3,
    });
    generatedSlide.addText('No bullet level two', {
      x: 10,
      y: 7.5,
      w: 2,
      h: 0.5,
      indentLevel: 2,
    });
    generatedSlide.addText('Left\tCenter\tRight\tDecimal', {
      x: 1,
      y: 7.25,
      w: 8,
      h: 0.5,
      tabStops: [
        { position: 1 },
        { position: 2.25, alignment: 'ctr' },
        { position: 3.5, alignment: 'r' },
        { position: 4.75, alignment: 'dec' },
      ],
    });
    generatedSlide.addText('Empty tabs', {
      x: 1,
      y: 7.75,
      w: 3,
      h: 0.5,
      tabStops: [],
    });
    generatedSlide.addText(
      [
        { text: 'First\tA', options: { breakLine: true, tabStops: [{ position: 1.5, alignment: 'r' }] } },
        { text: 'Second\tB', options: { tabStops: [{ position: 2.5, alignment: 'ctr' }] } },
      ],
      { x: 5, y: 7.5, w: 4, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Single', options: { underline: true } },
        { text: ' Double', options: { underline: { style: 'dbl', color: 'ff0000' } } },
        { text: ' Wavy', options: { underline: { style: 'wavyDbl' } } },
        { text: ' None', options: { underline: { style: 'none' } } },
        { text: ' Dot dash', options: { underline: { style: 'dotDashHeavy' } } },
      ],
      { x: 9, y: 3, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'True', options: { strike: true } },
        { text: ' False', options: { strike: false } },
        { text: ' Single', options: { strike: 'sngStrike' } },
        { text: ' Double', options: { strike: 'dblStrike' } },
        { text: ' None', options: { strike: 'noStrike' } },
      ],
      { x: 9, y: 4, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Yellow', options: { highlight: 'ffff00' } },
        { text: ' Theme', options: { highlight: 'accent2' } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 5, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Red', options: { outline: { color: 'ff0000', size: 1.5 } } },
        { text: ' Theme', options: { outline: { color: 'accent1', size: 2 } } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 6, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Red', options: { glow: { color: 'ff0000', size: 8, opacity: 0.5 } } },
        { text: ' Theme', options: { glow: { color: 'accent1', size: 2.5, opacity: 1 } } },
        { text: ' Default', options: { glow: { size: 0, opacity: 0 } } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 7, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Super', options: { superscript: true } },
        { text: ' Sub', options: { subscript: true } },
        { text: ' Custom+', options: { baseline: 600 } },
        { text: ' Custom-', options: { baseline: -800 } },
        { text: ' Fraction', options: { baseline: 1.5 } },
        { text: ' Zero', options: { baseline: 0 } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 8, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Positive', options: { charSpacing: 2.5 } },
        { text: ' Negative', options: { charSpacing: -1.25 } },
        { text: ' Fraction', options: { charSpacing: 0.004 } },
        { text: ' Zero', options: { charSpacing: 0 } },
        { text: ' Combined', options: { charSpacing: 3, baseline: 600 } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 9, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Inherited', options: {} },
        { text: ' German', options: { lang: 'de-DE' } },
        { text: ' Explicit default', options: { lang: 'en-US' } },
        { text: ' Empty inherits', options: { lang: '' } },
      ],
      { x: 9, y: 10, w: 3, h: 1, lang: 'fr-CA', objectName: 'Language outer' },
    );
    generatedSlide.addText('Omitted margin', {
      x: 0,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Margin omitted',
    });
    generatedSlide.addText('Zero margin', {
      x: 0,
      y: 0.5,
      w: 2,
      h: 0.5,
      margin: 0,
      objectName: 'Margin zero',
    });
    generatedSlide.addText('Scalar margin', {
      x: 0,
      y: 1,
      w: 2,
      h: 0.5,
      margin: 10,
      objectName: 'Margin scalar',
    });
    generatedSlide.addText('Tuple margin', {
      x: 0,
      y: 1.5,
      w: 2,
      h: 0.5,
      margin: [4, 8, 8, 4],
      objectName: 'Margin tuple',
    });
    generatedSlide.addText('Fractional margin', {
      x: 0,
      y: 2,
      w: 2,
      h: 0.5,
      margin: 0.125,
      objectName: 'Margin fractional',
    });
    generatedSlide.addText('Negative margin', {
      x: 0,
      y: 2.5,
      w: 2,
      h: 0.5,
      margin: -0.5,
      objectName: 'Margin negative',
    });
    generatedSlide.addText('Asymmetric probe', {
      x: 0,
      y: 3,
      w: 2,
      h: 0.5,
      margin: [1, 2, 3, 4],
      objectName: 'Margin asymmetric probe',
    });
    generatedSlide.addText('Omitted vertical alignment', {
      x: 2,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Vertical omitted',
    });
    generatedSlide.addText('Top vertical alignment', {
      x: 2,
      y: 0.5,
      w: 2,
      h: 0.5,
      valign: 'top',
      objectName: 'Vertical top',
    });
    generatedSlide.addText('Middle vertical alignment', {
      x: 2,
      y: 1,
      w: 2,
      h: 0.5,
      valign: 'middle',
      objectName: 'Vertical middle',
    });
    generatedSlide.addText('Bottom vertical alignment', {
      x: 2,
      y: 1.5,
      w: 2,
      h: 0.5,
      valign: 'bottom',
      objectName: 'Vertical bottom',
    });
    generatedSlide.addText(
      [{ text: 'Ignored run vertical alignment', options: { valign: 'bottom' } }],
      {
        x: 2,
        y: 2,
        w: 2,
        h: 0.5,
        objectName: 'Vertical run ignored',
      },
    );
    generatedSlide.addText('Omitted text wrapping', {
      x: 4,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Wrap omitted',
    });
    generatedSlide.addText('Enabled text wrapping', {
      x: 4,
      y: 0.5,
      w: 2,
      h: 0.5,
      wrap: true,
      objectName: 'Wrap true',
    });
    generatedSlide.addText('Disabled text wrapping', {
      x: 4,
      y: 1,
      w: 2,
      h: 0.5,
      wrap: false,
      objectName: 'Wrap false',
    });
    generatedSlide.addText('Invalid text wrapping', {
      x: 4,
      y: 1.5,
      w: 2,
      h: 0.5,
      wrap: 'false',
      objectName: 'Wrap invalid fallback',
    });
    generatedSlide.addText(
      [{ text: 'Ignored run text wrapping', options: { wrap: false } }],
      {
        x: 4,
        y: 2,
        w: 2,
        h: 0.5,
        objectName: 'Wrap run ignored',
      },
    );
    const textDirections = [
      'eaVert',
      'horz',
      'mongolianVert',
      'vert',
      'vert270',
      'wordArtVert',
      'wordArtVertRtl',
    ] as const;
    generatedSlide.addText('Omitted text direction', {
      x: 6,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Direction omitted',
    });
    for (const [index, direction] of textDirections.entries()) {
      generatedSlide.addText(direction, {
        x: 6,
        y: (index + 1) * 0.5,
        w: 2,
        h: 0.5,
        vert: direction,
        objectName: `Direction ${direction}`,
      });
    }
    generatedSlide.addText('Invalid text direction', {
      x: 6,
      y: 4,
      w: 2,
      h: 0.5,
      vert: 'vertical',
      objectName: 'Direction invalid passthrough',
    });
    generatedSlide.addText('Ignored textDirection alias', {
      x: 6,
      y: 4.5,
      w: 2,
      h: 0.5,
      textDirection: 'vert',
      objectName: 'Direction alias ignored',
    });
    generatedSlide.addText(
      [{ text: 'Ignored run text direction', options: { vert: 'vert270', textDirection: 'vert' } }],
      {
        x: 6,
        y: 5,
        w: 2,
        h: 0.5,
        objectName: 'Direction run ignored',
      },
    );
    generatedSlide.addText('Omitted text fit', {
      x: 8,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Fit omitted',
    });
    generatedSlide.addText('No text fit', {
      x: 8,
      y: 0.5,
      w: 2,
      h: 0.5,
      fit: 'none',
      objectName: 'Fit none',
    });
    generatedSlide.addText('Shrink text fit', {
      x: 8,
      y: 1,
      w: 2,
      h: 0.5,
      fit: 'shrink',
      objectName: 'Fit shrink',
    });
    generatedSlide.addText('Resize text fit', {
      x: 8,
      y: 1.5,
      w: 2,
      h: 0.5,
      fit: 'resize',
      objectName: 'Fit resize',
    });
    generatedSlide.addText('Invalid text fit', {
      x: 8,
      y: 2,
      w: 2,
      h: 0.5,
      fit: 'SHRINK',
      objectName: 'Fit invalid ignored',
    });
    generatedSlide.addText('Legacy shrink text fit', {
      x: 8,
      y: 2.5,
      w: 2,
      h: 0.5,
      shrinkText: true,
      objectName: 'Fit legacy shrink',
    });
    generatedSlide.addText('Legacy resize text fit', {
      x: 8,
      y: 3,
      w: 2,
      h: 0.5,
      autoFit: true,
      objectName: 'Fit legacy resize',
    });
    generatedSlide.addText(
      [{ text: 'Ignored run text fit', options: { fit: 'shrink', shrinkText: true, autoFit: true } }],
      {
        x: 8,
        y: 3.5,
        w: 2,
        h: 0.5,
        objectName: 'Fit run ignored',
      },
    );
    generatedSlide.addText('مرحبا\nالعالم', {
      x: 0,
      y: 4,
      w: 2,
      h: 1,
      objectName: 'RTL true',
      rtlMode: true,
    });
    generatedSlide.addText('Explicit false', {
      x: 0,
      y: 5,
      w: 2,
      h: 0.5,
      objectName: 'RTL false',
      rtlMode: false,
    });
    generatedSlide.addText('Omitted', {
      x: 0,
      y: 5.5,
      w: 2,
      h: 0.5,
      objectName: 'RTL omitted',
    });
    generatedSlide.addText(
      [
        { text: 'Run one', options: { rtlMode: true } },
        { text: ' Run two', options: { rtlMode: true } },
      ],
      {
        x: 0,
        y: 6,
        w: 2,
        h: 0.5,
        objectName: 'RTL run probe',
      },
    );
    generatedSlide.addText('Inset one inch', {
      x: 2,
      y: 4,
      w: 2,
      h: 0.5,
      inset: 1,
      objectName: 'Inset one',
    });
    generatedSlide.addText('Inset explicit zero', {
      x: 2,
      y: 4.5,
      w: 2,
      h: 0.5,
      inset: 0,
      objectName: 'Inset zero',
    });
    generatedSlide.addText('Equivalent modern margin', {
      x: 2,
      y: 5,
      w: 2,
      h: 0.5,
      margin: 72,
      objectName: 'Inset margin equivalent',
    });
    generatedSlide.addText('Modern margin precedence', {
      x: 2,
      y: 5.5,
      w: 2,
      h: 0.5,
      inset: 1,
      margin: 10,
      objectName: 'Inset modern margin wins',
    });
    generatedSlide.addText(
      [{ text: 'Run inset is inert', options: { inset: 1 } }],
      {
        x: 2,
        y: 6,
        w: 2,
        h: 0.5,
        objectName: 'Inset run ignored',
      },
    );
    const document = await importPptxGenJS(generated);
    expect(document.slides[0]?.title.text).toBe('Created by PptxGenJS');
    expect((document.slides[0]!.shapes[0] as ShapeModel).richText[0]!.align).toBe('center');
    const rich = document.slides[0]!.shapes[1] as ShapeModel;
    expect(rich.text).toBe('Bold red\nitalic');
    expect(rich.richText[0]!.runs).toEqual([
      {
        text: 'Bold red',
        style: {
          fontFamily: 'Aptos',
          fontSize: 24,
          lang: 'en-US',
          bold: true,
          color: { kind: 'srgb', value: 'FF0000' },
        },
      },
      {
        text: 'italic',
        softBreakBefore: true,
        style: {
          fontSize: 14,
          lang: 'en-US',
          italic: true,
          color: { kind: 'srgb', value: '4472C4' },
        },
      },
    ]);
    const aligned = document.slides[0]!.shapes[2] as ShapeModel;
    expect(aligned.richText.map(({ align }) => align)).toEqual(['left', 'center', 'right', 'justify']);
    aligned.richText = aligned.richText.map((paragraph, index) => ({
      runs: paragraph.runs,
      ...(index === 3
        ? { align: 'center' as const }
        : paragraph.align
          ? { align: paragraph.align }
          : {}),
    }));
    expect((document.slides[0]!.shapes[3] as ShapeModel).richText.map(({ bullet }) => bullet)).toEqual([
      { kind: 'bullet', character: '•', indent: 27 },
      { kind: 'bullet', character: '•', indent: 27 },
    ]);
    expect((document.slides[0]!.shapes[4] as ShapeModel).richText[0]!.bullet).toEqual({
      kind: 'bullet',
      character: '►',
      indent: 18,
    });
    expect((document.slides[0]!.shapes[5] as ShapeModel).richText[0]!.bullet).toEqual({
      kind: 'number',
      style: 'arabicPeriod',
      startAt: 3,
      indent: 22,
    });
    expect((document.slides[0]!.shapes[6] as ShapeModel).richText[0]!.bullet).toEqual({
      kind: 'number',
      style: 'romanLcParenR',
      startAt: 4,
      indent: 24,
    });
    expect((document.slides[0]!.shapes[7] as ShapeModel).richText.map(({ spacing }) => spacing)).toEqual([
      { before: 6.25, after: 8.5, line: { kind: 'exact', points: 28 } },
      { before: 6.25, after: 8.5, line: { kind: 'exact', points: 28 } },
    ]);
    expect((document.slides[0]!.shapes[8] as ShapeModel).richText[0]!.spacing).toEqual({
      before: 4.25,
      after: 7.75,
      line: { kind: 'multiple', factor: 1.5 },
    });
    expect((document.slides[0]!.shapes[9] as ShapeModel).richText[0]!.spacing).toBeUndefined();
    expect((document.slides[0]!.shapes[10] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'bullet', character: '•', indent: 27 },
      level: 1,
    });
    expect((document.slides[0]!.shapes[11] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'bullet', character: '►', indent: 18 },
      level: 2,
    });
    expect((document.slides[0]!.shapes[12] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
      level: 3,
    });
    const noBulletLevel = (document.slides[0]!.shapes[13] as ShapeModel).richText[0]!;
    expect(noBulletLevel.level).toBe(2);
    expect(noBulletLevel.bullet).toBeUndefined();
    expect((document.slides[0]!.shapes[14] as ShapeModel).richText[0]!.tabStops).toEqual([
      { position: 1, alignment: 'left' },
      { position: 2.25, alignment: 'center' },
      { position: 3.5, alignment: 'right' },
      { position: 4.75, alignment: 'decimal' },
    ]);
    expect((document.slides[0]!.shapes[15] as ShapeModel).richText[0]!.tabStops).toEqual([]);
    expect((document.slides[0]!.shapes[16] as ShapeModel).richText.map(({ tabStops }) => tabStops)).toEqual([
      [{ position: 1.5, alignment: 'right' }],
      [{ position: 2.5, alignment: 'center' }],
    ]);
    expect((document.slides[0]!.shapes[17] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.underline,
    )).toEqual([
      { style: 'sng' },
      { style: 'dbl', color: { kind: 'srgb', value: 'FF0000' } },
      { style: 'wavyDbl' },
      false,
      { style: 'dotDashHeavy' },
    ]);
    expect((document.slides[0]!.shapes[18] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.strike,
    )).toEqual(['sngStrike', undefined, 'sngStrike', 'dblStrike', false]);
    expect((document.slides[0]!.shapes[19] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.highlight,
    )).toEqual([
      { kind: 'srgb', value: 'FFFF00' },
      { kind: 'scheme', value: 'accent2' },
      undefined,
    ]);
    expect((document.slides[0]!.shapes[20] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.outline,
    )).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, size: 1.5 },
      { color: { kind: 'scheme', value: 'accent1' }, size: 2 },
      undefined,
    ]);
    expect((document.slides[0]!.shapes[21] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.glow,
    )).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, opacity: 0.5, size: 8 },
      { color: { kind: 'scheme', value: 'accent1' }, opacity: 1, size: 2.5 },
      { color: { kind: 'srgb', value: 'FFFFFF' }, opacity: 0, size: 0 },
      undefined,
    ]);
    expect((document.slides[0]!.shapes[22] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.baseline,
    )).toEqual(['superscript', 'subscript', 'superscript', 'subscript', 0.075, undefined, undefined]);
    const spaced = (document.slides[0]!.shapes[23] as ShapeModel).richText[0]!.runs;
    expect(spaced.map(({ style }) => style?.characterSpacing)).toEqual([2.5, -1.25, 0, undefined, 3, undefined]);
    expect(spaced[4]!.style!.baseline).toBe('superscript');
    const languages = (document.slides[0]!.shapes[24] as ShapeModel).richText[0]!.runs;
    expect(languages.map(({ style }) => style?.lang)).toEqual([
      'fr-CA',
      'de-DE',
      'en-US',
      'fr-CA',
    ]);
    const shapeByName = (name: string): ShapeModel => {
      const shape = document.slides[0]!.shapes.find((candidate) => candidate.name === name);
      expect(shape).toBeInstanceOf(ShapeModel);
      return shape as ShapeModel;
    };
    expect(shapeByName('RTL true').richText.map(({ rtl }) => rtl)).toEqual([true, true]);
    expect(shapeByName('RTL false').richText.map(({ rtl }) => rtl)).toEqual([undefined]);
    expect(shapeByName('RTL omitted').richText.map(({ rtl }) => rtl)).toEqual([undefined]);
    const rtlRunProbe = shapeByName('RTL run probe');
    expect(rtlRunProbe.text).toBe('Run one Run two');
    expect(rtlRunProbe.richText[0]!.runs.map(({ style }) =>
      (style as Record<string, unknown> | undefined)?.rtlMode)).toEqual([undefined, undefined]);
    expect(shapeByName('Margin omitted').textMargins).toBeUndefined();
    expect(shapeByName('Margin zero').textMargins).toEqual({ top: 0, right: 0, bottom: 0, left: 0 });
    expect(shapeByName('Margin scalar').textMargins).toEqual({
      top: 10,
      right: 10,
      bottom: 10,
      left: 10,
    });
    expect(shapeByName('Margin tuple').textMargins).toEqual({ top: 4, right: 8, bottom: 8, left: 4 });
    expect(shapeByName('Margin fractional').textMargins).toEqual({
      top: 1_588 / 12_700,
      right: 1_588 / 12_700,
      bottom: 1_588 / 12_700,
      left: 1_588 / 12_700,
    });
    expect(shapeByName('Margin negative').textMargins).toEqual({
      top: -0.5,
      right: -0.5,
      bottom: -0.5,
      left: -0.5,
    });
    expect(shapeByName('Margin asymmetric probe').textMargins).toEqual({
      top: 4,
      right: 2,
      bottom: 3,
      left: 1,
    });
    expect([
      'Inset one',
      'Inset zero',
      'Inset margin equivalent',
      'Inset modern margin wins',
      'Inset run ignored',
    ].map((name) => shapeByName(name).textMargins)).toEqual([
      { top: 72, right: 72, bottom: 72, left: 72 },
      { top: 0, right: 0, bottom: 0, left: 0 },
      { top: 72, right: 72, bottom: 72, left: 72 },
      { top: 10, right: 10, bottom: 10, left: 10 },
      undefined,
    ]);
    expect([
      'Vertical omitted',
      'Vertical top',
      'Vertical middle',
      'Vertical bottom',
      'Vertical run ignored',
    ].map((name) => shapeByName(name).verticalAlignment)).toEqual([
      'middle',
      'top',
      'middle',
      'bottom',
      'middle',
    ]);
    expect([
      'Wrap omitted',
      'Wrap true',
      'Wrap false',
      'Wrap invalid fallback',
      'Wrap run ignored',
    ].map((name) => shapeByName(name).textWrap)).toEqual([
      true,
      true,
      false,
      true,
      true,
    ]);
    expect([
      'Direction omitted',
      ...textDirections.map((direction) => `Direction ${direction}`),
      'Direction invalid passthrough',
      'Direction alias ignored',
      'Direction run ignored',
    ].map((name) => shapeByName(name).textDirection)).toEqual([
      undefined,
      ...textDirections,
      undefined,
      undefined,
      undefined,
    ]);
    expect([
      'Fit omitted',
      'Fit none',
      'Fit shrink',
      'Fit resize',
      'Fit invalid ignored',
      'Fit legacy shrink',
      'Fit legacy resize',
      'Fit run ignored',
    ].map((name) => shapeByName(name).textFit)).toEqual([
      undefined,
      undefined,
      'shrink',
      'resize',
      undefined,
      'shrink',
      'resize',
      undefined,
    ]);
    const importedXml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    const expectInsetAttributes = (name: string, value: number) => {
      expect(importedXml).toMatch(new RegExp(
        `name="${name}"[\\s\\S]*?<a:bodyPr`
          + `(?=[^>]*lIns="${value}")`
          + `(?=[^>]*tIns="${value}")`
          + `(?=[^>]*rIns="${value}")`
          + `(?=[^>]*bIns="${value}")[^>]*>`,
      ));
    };
    expectInsetAttributes('Inset one', 914400);
    expectInsetAttributes('Inset zero', 0);
    expectInsetAttributes('Inset margin equivalent', 914400);
    expectInsetAttributes('Inset modern margin wins', 127000);
    expect(importedXml).toMatch(
      /name="Inset run ignored"[\s\S]*?<a:bodyPr(?![^>]*(?:lIns|tIns|rIns|bIns)=)[^>]*>/,
    );
    expect(importedXml).toContain('lIns="12700" tIns="50800" rIns="25400" bIns="38100"');
    expect(importedXml).toMatch(/name="Vertical omitted"[\s\S]*?<a:bodyPr[^>]*anchor="ctr"/);
    expect(importedXml).toMatch(/name="Vertical top"[\s\S]*?<a:bodyPr[^>]*anchor="t"/);
    expect(importedXml).toMatch(/name="Vertical middle"[\s\S]*?<a:bodyPr[^>]*anchor="ctr"/);
    expect(importedXml).toMatch(/name="Vertical bottom"[\s\S]*?<a:bodyPr[^>]*anchor="b"/);
    expect(importedXml).toMatch(/name="Vertical run ignored"[\s\S]*?<a:bodyPr[^>]*anchor="ctr"/);
    expect(importedXml).toMatch(/name="Wrap omitted"[\s\S]*?<a:bodyPr[^>]*wrap="square"/);
    expect(importedXml).toMatch(/name="Wrap true"[\s\S]*?<a:bodyPr[^>]*wrap="square"/);
    expect(importedXml).toMatch(/name="Wrap false"[\s\S]*?<a:bodyPr[^>]*wrap="none"/);
    expect(importedXml).toMatch(/name="Wrap invalid fallback"[\s\S]*?<a:bodyPr[^>]*wrap="square"/);
    expect(importedXml).toMatch(/name="Wrap run ignored"[\s\S]*?<a:bodyPr[^>]*wrap="square"/);
    expect(importedXml).toMatch(
      /name="Direction omitted"[\s\S]*?<a:bodyPr(?![^>]*\svert=)[^>]*>/,
    );
    for (const direction of textDirections) {
      expect(importedXml).toMatch(
        new RegExp(`name="Direction ${direction}"[\\s\\S]*?<a:bodyPr[^>]* vert="${direction}"`),
      );
    }
    expect(importedXml).toMatch(
      /name="Direction invalid passthrough"[\s\S]*?<a:bodyPr[^>]* vert="vertical"/,
    );
    expect(importedXml).toMatch(
      /name="Direction alias ignored"[\s\S]*?<a:bodyPr(?![^>]*\svert=)[^>]*>/,
    );
    expect(importedXml).toMatch(
      /name="Direction run ignored"[\s\S]*?<a:bodyPr(?![^>]*\svert=)[^>]*>/,
    );
    expect(importedXml).toMatch(
      /name="Fit omitted"[\s\S]*?<a:bodyPr[^>]*(?:\/>|><\/a:bodyPr>)/,
    );
    expect(importedXml).toMatch(
      /name="Fit none"[\s\S]*?<a:bodyPr[^>]*(?:\/>|><\/a:bodyPr>)/,
    );
    expect(importedXml).toMatch(
      /name="Fit shrink"[\s\S]*?<a:bodyPr[^>]*><a:normAutofit\/><\/a:bodyPr>/,
    );
    expect(importedXml).toMatch(
      /name="Fit resize"[\s\S]*?<a:bodyPr[^>]*><a:spAutoFit\/><\/a:bodyPr>/,
    );
    expect(importedXml).toMatch(
      /name="Fit invalid ignored"[\s\S]*?<a:bodyPr[^>]*(?:\/>|><\/a:bodyPr>)/,
    );
    expect(importedXml).toMatch(
      /name="Fit legacy shrink"[\s\S]*?<a:bodyPr[^>]*><a:normAutofit\/><\/a:bodyPr>/,
    );
    expect(importedXml).toMatch(
      /name="Fit legacy resize"[\s\S]*?<a:bodyPr[^>]*><a:spAutoFit\/><\/a:bodyPr>/,
    );
    expect(importedXml).toMatch(
      /name="Fit run ignored"[\s\S]*?<a:bodyPr[^>]*(?:\/>|><\/a:bodyPr>)/,
    );
    expect(importedXml).toMatch(
      /name="Language outer"[\s\S]*?<a:rPr lang="fr-CA" altLang="en-US" dirty="0">/,
    );
    expect(importedXml).toMatch(
      /name="Language outer"[\s\S]*?<a:rPr lang="de-DE" altLang="en-US" dirty="0">/,
    );
    expect(importedXml).toMatch(
      /name="Language outer"[\s\S]*?<a:endParaRPr lang="fr-CA" dirty="0"\/>/,
    );
    expect(importedXml).toMatch(
      /name="RTL true"[\s\S]*?<a:p><a:pPr rtl="1"[^>]*>[\s\S]*?<a:p><a:pPr rtl="1"/,
    );
    expect(importedXml).toMatch(
      /name="RTL false"[\s\S]*?<a:p><a:pPr(?![^>]*\srtl=)[^>]*>/,
    );
    expect(importedXml).toMatch(
      /name="RTL omitted"[\s\S]*?<a:p><a:pPr(?![^>]*\srtl=)[^>]*>/,
    );
    const rtlRunStart = importedXml.indexOf('name="RTL run probe"');
    const rtlRunEnd = importedXml.indexOf('</p:sp>', rtlRunStart);
    expect(importedXml.slice(rtlRunStart, rtlRunEnd).match(/<a:pPr rtl="1"/g)).toHaveLength(2);
    document.slides[0]!.title.text = 'Edited by the OOXML kernel';
    document.duplicateSlide(0);

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides.map(({ title }) => title.text)).toEqual([
      'Edited by the OOXML kernel',
      'Edited by the OOXML kernel',
    ]);
    expect((reopened.slides[1]!.shapes[1] as ShapeModel).richText[0]!.runs[1]).toMatchObject({
      text: 'italic',
      softBreakBefore: true,
      style: { italic: true, color: { kind: 'srgb', value: '4472C4' } },
    });
    expect((reopened.slides[1]!.shapes[2] as ShapeModel).richText.map(({ align }) => align)).toEqual([
      'left',
      'center',
      'right',
      'center',
    ]);
    expect((reopened.slides[1]!.shapes[4] as ShapeModel).richText[0]!.bullet).toEqual({
      kind: 'bullet',
      character: '►',
      indent: 18,
    });
    expect((reopened.slides[1]!.shapes[6] as ShapeModel).richText[0]!.bullet).toMatchObject({
      kind: 'number',
      style: 'romanLcParenR',
      startAt: 4,
    });
    expect((reopened.slides[1]!.shapes[7] as ShapeModel).richText[0]!.spacing).toEqual({
      before: 6.25,
      after: 8.5,
      line: { kind: 'exact', points: 28 },
    });
    expect((reopened.slides[1]!.shapes[8] as ShapeModel).richText[0]!.spacing).toEqual({
      before: 4.25,
      after: 7.75,
      line: { kind: 'multiple', factor: 1.5 },
    });
    expect((reopened.slides[1]!.shapes[10] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'bullet', indent: 27 },
      level: 1,
    });
    expect((reopened.slides[1]!.shapes[12] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'number', indent: 22 },
      level: 3,
    });
    expect((reopened.slides[1]!.shapes[14] as ShapeModel).richText[0]!.tabStops).toEqual([
      { position: 1, alignment: 'left' },
      { position: 2.25, alignment: 'center' },
      { position: 3.5, alignment: 'right' },
      { position: 4.75, alignment: 'decimal' },
    ]);
    expect((reopened.slides[1]!.shapes[16] as ShapeModel).richText.map(({ tabStops }) => tabStops)).toEqual([
      [{ position: 1.5, alignment: 'right' }],
      [{ position: 2.5, alignment: 'center' }],
    ]);
    expect((reopened.slides[1]!.shapes[17] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.underline,
    )).toEqual([
      { style: 'sng' },
      { style: 'dbl', color: { kind: 'srgb', value: 'FF0000' } },
      { style: 'wavyDbl' },
      false,
      { style: 'dotDashHeavy' },
    ]);
    expect((reopened.slides[1]!.shapes[18] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.strike,
    )).toEqual(['sngStrike', undefined, 'sngStrike', 'dblStrike', false]);
    expect((reopened.slides[1]!.shapes[19] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.highlight,
    )).toEqual([
      { kind: 'srgb', value: 'FFFF00' },
      { kind: 'scheme', value: 'accent2' },
      undefined,
    ]);
    expect((reopened.slides[1]!.shapes[20] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.outline,
    )).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, size: 1.5 },
      { color: { kind: 'scheme', value: 'accent1' }, size: 2 },
      undefined,
    ]);
    expect((reopened.slides[1]!.shapes[21] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.glow,
    )).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, opacity: 0.5, size: 8 },
      { color: { kind: 'scheme', value: 'accent1' }, opacity: 1, size: 2.5 },
      { color: { kind: 'srgb', value: 'FFFFFF' }, opacity: 0, size: 0 },
      undefined,
    ]);
    expect((reopened.slides[1]!.shapes[22] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.baseline,
    )).toEqual(['superscript', 'subscript', 'superscript', 'subscript', 0.075, undefined, undefined]);
    const reopenedSpaced = (reopened.slides[1]!.shapes[23] as ShapeModel).richText[0]!.runs;
    expect(reopenedSpaced.map(({ style }) => style?.characterSpacing))
      .toEqual([2.5, -1.25, 0, undefined, 3, undefined]);
    expect(reopenedSpaced[4]!.style!.baseline).toBe('superscript');
    const reopenedLanguages = (reopened.slides[1]!.shapes[24] as ShapeModel).richText[0]!.runs;
    expect(reopenedLanguages.map(({ style }) => style?.lang)).toEqual([
      'fr-CA',
      'de-DE',
      'en-US',
      'fr-CA',
    ]);
    const reopenedMargins = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Margin '))
      .map(({ name, textMargins }) => [name, textMargins]);
    expect(reopenedMargins).toEqual([
      ['Margin omitted', undefined],
      ['Margin zero', { top: 0, right: 0, bottom: 0, left: 0 }],
      ['Margin scalar', { top: 10, right: 10, bottom: 10, left: 10 }],
      ['Margin tuple', { top: 4, right: 8, bottom: 8, left: 4 }],
      ['Margin fractional', {
        top: 1_588 / 12_700,
        right: 1_588 / 12_700,
        bottom: 1_588 / 12_700,
        left: 1_588 / 12_700,
      }],
      ['Margin negative', { top: -0.5, right: -0.5, bottom: -0.5, left: -0.5 }],
      ['Margin asymmetric probe', { top: 4, right: 2, bottom: 3, left: 1 }],
    ]);
    const reopenedInsets = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Inset '))
      .map(({ name, textMargins }) => [name, textMargins]);
    expect(reopenedInsets).toEqual([
      ['Inset one', { top: 72, right: 72, bottom: 72, left: 72 }],
      ['Inset zero', { top: 0, right: 0, bottom: 0, left: 0 }],
      ['Inset margin equivalent', { top: 72, right: 72, bottom: 72, left: 72 }],
      ['Inset modern margin wins', { top: 10, right: 10, bottom: 10, left: 10 }],
      ['Inset run ignored', undefined],
    ]);
    const reopenedVerticalAlignment = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Vertical '))
      .map(({ name, verticalAlignment }) => [name, verticalAlignment]);
    expect(reopenedVerticalAlignment).toEqual([
      ['Vertical omitted', 'middle'],
      ['Vertical top', 'top'],
      ['Vertical middle', 'middle'],
      ['Vertical bottom', 'bottom'],
      ['Vertical run ignored', 'middle'],
    ]);
    const reopenedWrapping = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Wrap '))
      .map(({ name, textWrap }) => [name, textWrap]);
    expect(reopenedWrapping).toEqual([
      ['Wrap omitted', true],
      ['Wrap true', true],
      ['Wrap false', false],
      ['Wrap invalid fallback', true],
      ['Wrap run ignored', true],
    ]);
    const reopenedDirections = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Direction '))
      .map(({ name, textDirection }) => [name, textDirection]);
    expect(reopenedDirections).toEqual([
      ['Direction omitted', undefined],
      ...textDirections.map((direction) => [`Direction ${direction}`, direction]),
      ['Direction invalid passthrough', undefined],
      ['Direction alias ignored', undefined],
      ['Direction run ignored', undefined],
    ]);
    const reopenedFit = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Fit '))
      .map(({ name, textFit }) => [name, textFit]);
    expect(reopenedFit).toEqual([
      ['Fit omitted', undefined],
      ['Fit none', undefined],
      ['Fit shrink', 'shrink'],
      ['Fit resize', 'resize'],
      ['Fit invalid ignored', undefined],
      ['Fit legacy shrink', 'shrink'],
      ['Fit legacy resize', 'resize'],
      ['Fit run ignored', undefined],
    ]);
    const reopenedRtl = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('RTL '))
      .map(({ name, richText }) => [name, richText.map(({ rtl }) => rtl)]);
    expect(reopenedRtl).toEqual([
      ['RTL true', [true, true]],
      ['RTL false', [undefined]],
      ['RTL omitted', [undefined]],
      ['RTL run probe', [true]],
    ]);
  }, 180_000);

  it('imports and reopens PptxGenJS rich text transparency from real output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    const slide = generated.addSlide();
    slide.addText(
      [
        { text: 'Omitted', options: { color: 'FF0000' } },
        { text: ' Zero', options: { color: '00FF00', transparency: 0 } },
        { text: ' Quarter', options: { color: '0000FF', transparency: 25 } },
        { text: ' Fractional', options: { color: '112233', transparency: 50.5555 } },
        { text: ' Invisible', options: { color: '445566', transparency: 100 } },
        { text: ' Theme', options: { color: 'accent1', transparency: 40 } },
        { text: ' Default', options: { transparency: 60 } },
      ],
      { x: 1, y: 1, w: 10, h: 1, objectName: 'Transparency probe' },
    );

    const document = await importPptxGenJS(generated);
    const shape = document.slides[0]!.shapes.find(({ name }) => name === 'Transparency probe');
    expect(shape).toBeInstanceOf(ShapeModel);
    const runs = (shape as ShapeModel).richText[0]!.runs;
    expect(runs.map(({ style }) => style?.transparency)).toEqual([
      undefined,
      undefined,
      25,
      50.555,
      100,
      40,
      60,
    ]);
    expect(runs.map(({ style }) => style?.color)).toEqual([
      { kind: 'srgb', value: 'FF0000' },
      { kind: 'srgb', value: '00FF00' },
      { kind: 'srgb', value: '0000FF' },
      { kind: 'srgb', value: '112233' },
      { kind: 'srgb', value: '445566' },
      { kind: 'scheme', value: 'accent1' },
      { kind: 'srgb', value: '000000' },
    ]);
    const slideXml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(slideXml.match(/<a:alpha val="\d+"\/>/g)).toEqual([
      '<a:alpha val="75000"/>',
      '<a:alpha val="49445"/>',
      '<a:alpha val="0"/>',
      '<a:alpha val="60000"/>',
      '<a:alpha val="40000"/>',
    ]);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedShape = reopened.slides[0]!.shapes.find(
      ({ name }) => name === 'Transparency probe',
    ) as ShapeModel;
    expect(reopenedShape.richText[0]!.runs.map(({ style }) => style?.transparency)).toEqual([
      undefined,
      undefined,
      25,
      50.555,
      100,
      40,
      60,
    ]);
  }, 30_000);

  it('imports only direct PptxGenJS presentation RTL and reopens it', async () => {
    const cases: readonly [string, boolean, unknown, boolean | undefined][] = [
      ['omitted', false, undefined, undefined],
      ['true', true, true, true],
      ['false', true, false, undefined],
      ['truthy', true, 'yes', true],
    ];
    for (const [name, assign, value, expected] of cases) {
      const generated = new PptxGenJS();
      if (assign) generated.rtlMode = value;
      generated.addSlide();
      const document = await importPptxGenJS(generated);
      const journal = [...document.opcPackage.mutations];

      expect(document.rtlMode, name).toBe(expected);
      expect(document.opcPackage.mutations, name).toEqual(journal);
      const presentationXml = new TextDecoder().decode(
        document.opcPackage.requirePart(document.presentationPartUri).bytes,
      );
      if (expected === true) {
        expect(presentationXml, name).toMatch(/<p:presentation\b[^>]*\srtl="1"/);
      } else {
        expect(presentationXml, name).not.toMatch(/<p:presentation\b[^>]*\srtl=/);
      }
      expect(presentationXml, name).toMatch(/<a:lvl1pPr\b[^>]*\srtl="0"/);

      if (name === 'true') {
        const reopened = await PptxDocument.open(await document.write());
        expect(reopened.rtlMode).toBe(true);
      }
    }
  }, 20_000);

  it('imports and matches public PptxGenJS hidden slide output', async () => {
    const cases: readonly (readonly [string, boolean, unknown, boolean])[] = [
      ['omitted', false, undefined, false],
      ['false', true, false, false],
      ['true', true, true, true],
      ['truthy invalid', true, 'yes', true],
    ];
    const generated = cases.map(([, assign, value]) => {
      const presentation = new PptxGenJS();
      const slide = presentation.addSlide();
      if (assign) slide.hidden = value;
      return presentation;
    });
    expect(generated.map(({ version }) => version)).toEqual(Array(4).fill('4.0.1'));

    const imported = await Promise.all(
      generated.map((presentation) => importPptxGenJS(presentation)),
    );
    expect(imported.map(({ slides }) => slides[0]?.hidden)).toEqual(
      cases.map(([, , , expected]) => expected),
    );
    for (const [index, document] of imported.entries()) {
      const [name, , , expected] = cases[index]!;
      const journal = [...document.opcPackage.mutations];
      const slide = document.slides[0]!;
      const slideXml = new TextDecoder().decode(
        document.opcPackage.requirePart(slide.partUri).bytes,
      );
      expect(document.slides[0]?.hidden, name).toBe(expected);
      expect(document.opcPackage.mutations, name).toEqual(journal);
      if (expected) {
        expect(slideXml, name).toMatch(/<p:sld\b[^>]*\sshow="0"/);
      } else {
        expect(slideXml, name).not.toMatch(/<p:sld\b[^>]*\sshow=/);
      }
      const reopened = await PptxDocument.open(await document.write());
      expect(reopened.slides[0]?.hidden, name).toBe(expected);
    }

    const native = PptxDocument.create();
    const nativeVisible = native.addSlide();
    const nativeHidden = native.addSlide();
    nativeHidden.hidden = true;
    expect([nativeVisible.hidden, nativeHidden.hidden]).toEqual([
      imported[1]!.slides[0]!.hidden,
      imported[2]!.slides[0]!.hidden,
    ]);
    const nativeXml = native.slides.map(({ partUri }) =>
      new TextDecoder().decode(native.opcPackage.requirePart(partUri).bytes));
    expect(nativeXml[0]).not.toMatch(/<p:sld\b[^>]*\sshow=/);
    expect(nativeXml[1]).toMatch(/<p:sld\b[^>]*\sshow="0"/);
    const reopenedNative = await PptxDocument.open(await native.write());
    expect(reopenedNative.slides.map(({ hidden }) => hidden)).toEqual([false, true]);

    const beforeInvalid = native.opcPackage.requirePart(nativeVisible.partUri).bytes.slice();
    const invalidJournal = [...native.opcPackage.mutations];
    expect(() => {
      (nativeVisible as unknown as { hidden: unknown }).hidden = 'yes';
    }).toThrow(TypeError);
    expect(native.opcPackage.requirePart(nativeVisible.partUri).bytes).toEqual(beforeInvalid);
    expect(native.opcPackage.mutations).toEqual(invalidJournal);
  }, 20_000);

  it('imports and matches public PptxGenJS speaker notes output', async () => {
    const cases = [
      { name: 'omitted', input: undefined, expected: '' },
      { name: 'empty', input: '', expected: '' },
      { name: 'plain', input: 'Speaker & <notes>', expected: 'Speaker & <notes>' },
      {
        name: 'multiline',
        input: 'Line 1\nLine 2\r\nLine 3',
        expected: 'Line 1\nLine 2\nLine 3',
      },
    ] as const;
    const generated = cases.map(({ input }) => {
      const presentation = new PptxGenJS();
      const slide = presentation.addSlide();
      if (input !== undefined) slide.addNotes(input);
      return presentation;
    });
    expect(generated.map(({ version }) => version)).toEqual(Array(4).fill('4.0.1'));

    const imported = await Promise.all(
      generated.map((presentation) => importPptxGenJS(presentation)),
    );
    expect(imported.map(({ slides }) => slides[0]?.notes)).toEqual(
      cases.map(({ expected }) => expected),
    );
    for (const [index, document] of imported.entries()) {
      const { name, expected } = cases[index]!;
      const slide = document.slides[0]!;
      const journal = [...document.opcPackage.mutations];
      expect(slide.notes, name).toBe(expected);
      expect(document.opcPackage.mutations, name).toEqual(journal);

      const notesRelationships = slide.relationships.filter(
        ({ type }) => type.endsWith('/notesSlide'),
      );
      expect(notesRelationships, name).toHaveLength(1);
      const notesUri = notesRelationships[0]!.resolvedTarget!;
      expect(document.opcPackage.requirePart(notesUri).contentType, name).toBe(
        'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
      );
      const slideBacklinks = document.opcPackage.relationships(notesUri).filter(
        ({ type }) => type.endsWith('/slide'),
      );
      const masterRelationships = document.opcPackage.relationships(notesUri).filter(
        ({ type }) => type.endsWith('/notesMaster'),
      );
      expect(slideBacklinks, name).toHaveLength(1);
      expect(slideBacklinks[0]!.resolvedTarget, name).toBe(slide.partUri);
      expect(masterRelationships, name).toHaveLength(1);
      const masterUri = masterRelationships[0]!.resolvedTarget!;
      expect(document.opcPackage.requirePart(masterUri).contentType, name).toBe(
        'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
      );
      const presentationMasterRelationships = document.opcPackage
        .relationships(document.presentationPartUri)
        .filter(({ type }) => type.endsWith('/notesMaster'));
      expect(presentationMasterRelationships, name).toHaveLength(1);
      expect(presentationMasterRelationships[0]!.resolvedTarget, name).toBe(masterUri);

      const notesXml = new TextDecoder().decode(
        document.opcPackage.requirePart(notesUri).bytes,
      );
      expect(notesXml.match(/<p:ph\b[^>]*\btype="body"/g), name).toHaveLength(1);
      expect(notesXml.match(/<p:txBody>/g), name).toHaveLength(2);
      if (name === 'plain') {
        expect(notesXml).toContain('<a:t>Speaker &amp; &lt;notes&gt;</a:t>');
      }

      const reopened = await PptxDocument.open(await document.write());
      expect(document.diagnostics.filter(({ severity }) => severity === 'error'), name)
        .toEqual([]);
      expect(reopened.slides[0]?.notes, name).toBe(expected);
    }

    const importedPlain = imported[2]!;
    const importedSource = importedPlain.slides[0]!;
    const importedSourceNotesUri = importedSource.relationships.find(
      ({ type }) => type.endsWith('/notesSlide'),
    )!.resolvedTarget!;
    const importedDuplicate = importedPlain.duplicateSlide(0);
    const importedDuplicateNotesUri = importedDuplicate.relationships.find(
      ({ type }) => type.endsWith('/notesSlide'),
    )!.resolvedTarget!;
    expect(importedDuplicate.notes).toBe('Speaker & <notes>');
    expect(importedDuplicateNotesUri).not.toBe(importedSourceNotesUri);
    expect(importedPlain.opcPackage.relationships(importedDuplicateNotesUri).find(
      ({ type }) => type.endsWith('/slide'),
    )?.resolvedTarget).toBe(importedDuplicate.partUri);
    importedDuplicate.notes = 'Edited duplicate';
    expect([importedSource.notes, importedDuplicate.notes]).toEqual([
      'Speaker & <notes>',
      'Edited duplicate',
    ]);
    const reopenedDuplicate = await PptxDocument.open(await importedPlain.write());
    expect(reopenedDuplicate.slides.map(({ notes }) => notes)).toEqual([
      'Speaker & <notes>',
      'Edited duplicate',
    ]);

    const native = [
      PptxDocument.create(),
      PptxDocument.create(),
      PptxDocument.create(),
      PptxDocument.create(),
    ];
    native[0]!.addSlide();
    native[1]!.addSlide().notes = '';
    native[2]!.addSlide().addNotes('Speaker & <notes>');
    native[3]!.addSlide().addNotes('Line 1\nLine 2\r\nLine 3');
    expect(native.map(({ slides }) => slides[0]?.notes)).toEqual([
      undefined,
      '',
      'Speaker & <notes>',
      'Line 1\nLine 2\nLine 3',
    ]);
    expect(native.slice(1).map(({ slides }) => slides[0]?.notes)).toEqual(
      imported.slice(1).map(({ slides }) => slides[0]?.notes),
    );
    const nativePlain = native[2]!.slides[0]!;
    const nativeNotesUri = nativePlain.relationships.find(
      ({ type }) => type.endsWith('/notesSlide'),
    )!.resolvedTarget!;
    expect(new TextDecoder().decode(
      native[2]!.opcPackage.requirePart(nativeNotesUri).bytes,
    )).toContain('<a:t xml:space="preserve">Speaker &amp; &lt;notes&gt;</a:t>');

    const invalidDocument = PptxDocument.create();
    const invalidSlide = invalidDocument.addSlide();
    const beforeParts = invalidDocument.opcPackage.parts.map(
      ({ uri, bytes }) => ({ uri, bytes: bytes.slice() }),
    );
    const beforeJournal = [...invalidDocument.opcPackage.mutations];
    for (const invalid of [7, {}, 'A\u0001B']) {
      expect(() => {
        (invalidSlide as unknown as { notes: unknown }).notes = invalid;
      }, String(invalid)).toThrow(TypeError);
      expect(() => invalidSlide.addNotes(invalid as never), String(invalid)).toThrow(TypeError);
    }
    expect(invalidDocument.opcPackage.parts.map(({ uri, bytes }) => ({ uri, bytes })))
      .toEqual(beforeParts);
    expect(invalidDocument.opcPackage.mutations).toEqual(beforeJournal);
  }, 20_000);

  it('imports and matches public PptxGenJS presentation sections', async () => {
    const none = new PptxGenJS();
    none.addSlide();
    none.addSlide();

    const explicit = new PptxGenJS();
    explicit.addSection({ title: 'Data & <One>' });
    explicit.addSlide({ sectionTitle: 'Data & <One>' });
    explicit.addSlide({ sectionTitle: 'Data & <One>' });

    const empty = new PptxGenJS();
    empty.addSlide();
    empty.addSection({ title: 'Empty' });

    const ordered = new PptxGenJS();
    ordered.addSection({ title: 'A' });
    ordered.addSection({ title: 'C' });
    ordered.addSection({ title: 'B', order: 1 });

    const defaults = new PptxGenJS();
    defaults.addSection({ title: 'Intro' });
    defaults.addSlide();
    defaults.addSlide();

    const looseBefore = new PptxGenJS();
    looseBefore.addSlide();
    looseBefore.addSection({ title: 'Later' });
    looseBefore.addSlide({ sectionTitle: 'Later' });

    const unknown = new PptxGenJS();
    unknown.addSection({ title: 'Known' });
    const warning = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
    try {
      unknown.addSlide({ sectionTitle: 'Missing' });
      expect(warning).toHaveBeenCalledTimes(1);
    } finally {
      warning.mockRestore();
    }

    const orderZero = new PptxGenJS();
    orderZero.addSection({ title: 'A' });
    orderZero.addSection({ title: 'Zero', order: 0 });

    const generated = [
      none,
      explicit,
      empty,
      ordered,
      defaults,
      looseBefore,
      unknown,
      orderZero,
    ];
    expect(generated.map(({ version }) => version)).toEqual(Array(8).fill('4.0.1'));
    const imported: PptxDocument[] = [];
    for (const presentation of generated) imported.push(await importPptxGenJS(presentation));

    expect(imported.map(sectionState)).toEqual([
      [],
      [{ title: 'Data & <One>', slideIds: [256, 257] }],
      [{ title: 'Empty', slideIds: [] }],
      [
        { title: 'A', slideIds: [] },
        { title: 'B', slideIds: [] },
        { title: 'C', slideIds: [] },
      ],
      [
        { title: 'Intro', slideIds: [] },
        { title: 'Default-1', slideIds: [256, 257] },
      ],
      [{ title: 'Later', slideIds: [257] }],
      [{ title: 'Known', slideIds: [] }],
      [
        { title: 'A', slideIds: [] },
        { title: 'Zero', slideIds: [] },
      ],
    ]);
    expect(imported.map(({ slides }) => slides.length)).toEqual([2, 2, 1, 0, 2, 2, 1, 0]);
    for (const [documentIndex, document] of imported.entries()) {
      for (const { id } of document.sections ?? []) {
        expect(id).toMatch(/^\{[0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$/i);
      }
      const masterless = documentIndex === 3 || documentIndex === 7;
      const output = await document.write(masterless ? { mode: 'permissive' } : {});
      expect(document.diagnostics.filter(({ code }) => code === 'LAYOUT_RELATIONSHIP_INVALID'))
        .toEqual(masterless
          ? [expect.objectContaining({
              severity: 'error',
              code: 'LAYOUT_RELATIONSHIP_INVALID',
              partUri: document.presentationPartUri,
              objectId: 'rId1',
            })]
          : []);
      const reopened = await PptxDocument.open(output);
      expect(sectionState(reopened)).toEqual(sectionState(document));
    }

    const nativeExplicit = PptxDocument.create();
    nativeExplicit.addSection({ title: 'Data & <One>' });
    nativeExplicit.addSlide({ sectionTitle: 'Data & <One>' });
    nativeExplicit.addSlide({ sectionTitle: 'Data & <One>' });
    expect(sectionState(nativeExplicit)).toEqual(sectionState(imported[1]!));

    const nativeOrdered = PptxDocument.create();
    nativeOrdered.addSection({ title: 'A' });
    nativeOrdered.addSection({ title: 'C' });
    nativeOrdered.addSection({ title: 'B', order: 1 });
    expect(sectionState(nativeOrdered)).toEqual(sectionState(imported[3]!));

    const nativeDefaults = PptxDocument.create();
    nativeDefaults.addSection({ title: 'Intro' });
    nativeDefaults.addSlide();
    nativeDefaults.addSlide();
    expect(sectionState(nativeDefaults)).toEqual(sectionState(imported[4]!));

    const nativeLooseBefore = PptxDocument.create();
    nativeLooseBefore.addSlide();
    nativeLooseBefore.addSection({ title: 'Later' });
    nativeLooseBefore.addSlide({ sectionTitle: 'Later' });
    expect(sectionState(nativeLooseBefore)).toEqual(sectionState(imported[5]!));

    const nativeUnknown = PptxDocument.create();
    nativeUnknown.addSection({ title: 'Known' });
    expect(() => nativeUnknown.addSlide({ sectionTitle: 'Missing' })).toThrow(RangeError);
    expect(nativeUnknown.slides).toHaveLength(0);
    expect(imported[6]!.slides).toHaveLength(1);

    const nativeOrderZero = PptxDocument.create();
    nativeOrderZero.addSection({ title: 'A' });
    nativeOrderZero.addSection({ title: 'Zero', order: 0 });
    expect(sectionState(nativeOrderZero)?.map(({ title }) => title)).toEqual(['Zero', 'A']);
    expect(sectionState(imported[7]!)?.map(({ title }) => title)).toEqual(['A', 'Zero']);
  }, 30_000);

  it('matches public PptxGenJS presentation theme fonts and reopens a partial edit', async () => {
    const cases = [
      {
        name: 'default',
        input: undefined,
        expected: { headFontFace: 'Calibri Light', bodyFontFace: 'Calibri' },
      },
      {
        name: 'empty',
        input: {},
        expected: { headFontFace: 'Calibri Light', bodyFontFace: 'Calibri' },
      },
      {
        name: 'head only',
        input: { headFontFace: 'Aptos Display' },
        expected: { headFontFace: 'Aptos Display', bodyFontFace: 'Calibri' },
      },
      {
        name: 'body only',
        input: { bodyFontFace: 'Aptos' },
        expected: { headFontFace: 'Calibri Light', bodyFontFace: 'Aptos' },
      },
      {
        name: 'custom',
        input: { headFontFace: 'Noto Sans Display', bodyFontFace: 'Noto Sans' },
        expected: { headFontFace: 'Noto Sans Display', bodyFontFace: 'Noto Sans' },
      },
    ] as const;
    const generated = cases.map(({ input }) => {
      const presentation = new PptxGenJS();
      if (input !== undefined) presentation.theme = input;
      presentation.addSlide();
      return presentation;
    });
    expect(generated.map(({ version }) => version)).toEqual(Array(5).fill('4.0.1'));

    const imported = await Promise.all(
      generated.map((presentation) => importPptxGenJS(presentation)),
    );
    for (const [index, document] of imported.entries()) {
      const journal = [...document.opcPackage.mutations];
      expect(document.theme, cases[index]!.name).toEqual(cases[index]!.expected);
      expect(document.opcPackage.mutations, cases[index]!.name).toEqual(journal);

      const native = PptxDocument.create({ theme: cases[index]!.input ?? {} });
      expect(native.theme, cases[index]!.name).toEqual(document.theme);
    }

    const edited = imported[4]!;
    const theme = edited.masterLayoutTheme.presentationTheme!;
    const untouchedParts = new Map(
      edited.opcPackage.parts
        .filter(({ uri }) => uri !== theme.partUri)
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    theme.setFonts({ minorLatin: 'Noto Sans Edited' });
    expect(edited.theme).toEqual({
      headFontFace: 'Noto Sans Display',
      bodyFontFace: 'Noto Sans Edited',
    });
    for (const [uri, bytes] of untouchedParts) {
      expect(edited.opcPackage.requirePart(uri).bytes).toEqual(bytes);
    }

    const reopened = await PptxDocument.open(await edited.write());
    expect(reopened.theme).toEqual(edited.theme);
  }, 30_000);

  it('imports, edits, and reopens PptxGenJS presentation created-at metadata from public output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.addSlide();
    const bytes = await generated.write({ outputType: 'uint8array', compression: true });
    const imported = await PptxDocument.open(bytes);
    const journal = [...imported.opcPackage.mutations];

    expect(imported.createdAt).toMatch(
      /^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}Z$/,
    );
    expect(imported.opcPackage.mutations).toEqual(journal);
    const createdAt = imported.createdAt!;
    const coreBefore = new TextDecoder().decode(
      imported.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    expect(coreBefore).toContain(
      'xmlns:dcterms="http://purl.org/dc/terms/"',
    );
    expect(coreBefore).toContain(
      'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"',
    );
    expect(coreBefore.match(
      /<dcterms:created xsi:type="dcterms:W3CDTF">([^<]+)<\/dcterms:created>/,
    )?.[1]).toBe(createdAt);
    const modifiedBefore = coreBefore.match(
      /<dcterms:modified xsi:type="dcterms:W3CDTF">[^<]+<\/dcterms:modified>/,
    )?.[0];
    expect(modifiedBefore).toMatch(
      /^<dcterms:modified xsi:type="dcterms:W3CDTF">[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}Z<\/dcterms:modified>$/,
    );
    const nonCoreParts = new Map(
      imported.opcPackage.parts
        .filter(({ uri }) => uri !== '/docProps/core.xml')
        .map(({ uri, bytes: partBytes }) => [uri, partBytes]),
    );

    const native = PptxDocument.create({ createdAt });
    expect(native.createdAt).toBe(createdAt);
    expect(PptxDocument.create().createdAt).toBeUndefined();

    imported.createdAt = '2024-02-29T12:34:56.123+05:30';
    expect(imported.createdAt).toBe('2024-02-29T12:34:56.123+05:30');
    const coreAfter = new TextDecoder().decode(
      imported.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    expect(coreAfter).not.toBe(coreBefore);
    expect(coreAfter.match(
      /<dcterms:modified xsi:type="dcterms:W3CDTF">[^<]+<\/dcterms:modified>/,
    )?.[0]).toBe(modifiedBefore);
    for (const [uri, partBytes] of nonCoreParts) {
      expect(imported.opcPackage.requirePart(uri).bytes).toEqual(partBytes);
    }

    const reopened = await PptxDocument.open(await imported.write());
    expect(reopened.createdAt).toBe('2024-02-29T12:34:56.123+05:30');
    const reopenedCore = new TextDecoder().decode(
      reopened.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    expect(reopenedCore.match(
      /<dcterms:modified xsi:type="dcterms:W3CDTF">[^<]+<\/dcterms:modified>/,
    )?.[0]).toBe(modifiedBefore);
  }, 20_000);

  it('imports, edits, and reopens PptxGenJS presentation modified-at metadata from public output', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.addSlide().addText('PptxGenJS modified-at', {
      x: 1,
      y: 1,
      w: 5,
      h: 1,
    });
    const imported = await importPptxGenJS(generated);
    const journal = [...imported.opcPackage.mutations];

    expect(imported.modifiedAt).toMatch(
      /^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}Z$/,
    );
    expect(imported.createdAt).toMatch(
      /^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}Z$/,
    );
    expect(imported.opcPackage.mutations).toEqual(journal);
    const modifiedAt = imported.modifiedAt!;
    const createdAt = imported.createdAt!;
    const coreBefore = new TextDecoder().decode(
      imported.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    expect(coreBefore.match(
      /<dcterms:modified xsi:type="dcterms:W3CDTF">([^<]+)<\/dcterms:modified>/,
    )?.[1]).toBe(modifiedAt);
    const createdBefore = coreBefore.match(
      /<dcterms:created xsi:type="dcterms:W3CDTF">[^<]+<\/dcterms:created>/,
    )?.[0];
    expect(createdBefore).toBe(
      `<dcterms:created xsi:type="dcterms:W3CDTF">${createdAt}</dcterms:created>`,
    );
    const nonCoreParts = new Map(
      imported.opcPackage.parts
        .filter(({ uri }) => uri !== '/docProps/core.xml')
        .map(({ uri, bytes }) => [uri, bytes]),
    );

    const native = PptxDocument.create({ modifiedAt });
    expect(native.modifiedAt).toBe(modifiedAt);
    expect(native.createdAt).toBeUndefined();
    expect(PptxDocument.create().modifiedAt).toBeUndefined();

    imported.modifiedAt = '2024-03-01T01:02:03.456+08:00';
    expect(imported.modifiedAt).toBe('2024-03-01T01:02:03.456+08:00');
    expect(imported.createdAt).toBe(createdAt);
    const coreAfter = new TextDecoder().decode(
      imported.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    expect(coreAfter).not.toBe(coreBefore);
    expect(coreAfter).toContain(
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2024-03-01T01:02:03.456+08:00</dcterms:modified>',
    );
    expect(coreAfter.match(
      /<dcterms:created xsi:type="dcterms:W3CDTF">[^<]+<\/dcterms:created>/,
    )?.[0]).toBe(createdBefore);
    for (const [uri, bytes] of nonCoreParts) {
      expect(imported.opcPackage.requirePart(uri).bytes).toEqual(bytes);
    }

    const reopened = await PptxDocument.open(await imported.write());
    expect(reopened.modifiedAt).toBe('2024-03-01T01:02:03.456+08:00');
    expect(reopened.createdAt).toBe(createdAt);
    expect(reopened.slides[0]?.title.text).toBe('PptxGenJS modified-at');
    const reopenedCore = new TextDecoder().decode(
      reopened.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    expect(reopenedCore.match(
      /<dcterms:created xsi:type="dcterms:W3CDTF">[^<]+<\/dcterms:created>/,
    )?.[0]).toBe(createdBefore);
  }, 20_000);

  it('imports and reopens PptxGenJS presentation title metadata from public output', async () => {
    const baseline = new PptxGenJS();
    baseline.addSlide();
    const custom = new PptxGenJS();
    custom.title = 'Quarterly & <Review>';
    custom.addSlide();
    const empty = new PptxGenJS();
    empty.title = '';
    empty.addSlide();
    expect([baseline.version, custom.version, empty.version]).toEqual([
      '4.0.1',
      '4.0.1',
      '4.0.1',
    ]);

    const expectedTitles = [
      'PptxGenJS Presentation',
      'Quarterly & <Review>',
      '',
    ] as const;
    const imported = await Promise.all([
      importPptxGenJS(baseline),
      importPptxGenJS(custom),
      importPptxGenJS(empty),
    ]);
    expect(imported.map(({ title }) => title)).toEqual(expectedTitles);
    for (const [index, document] of imported.entries()) {
      const journal = [...document.opcPackage.mutations];
      expect(document.title).toBe(expectedTitles[index]);
      expect(document.opcPackage.mutations).toEqual(journal);
    }

    const coreXml = imported.map((document) => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/core.xml').bytes,
    ));
    expect(coreXml[0]).toContain(
      '<dc:title>PptxGenJS Presentation</dc:title>',
    );
    expect(coreXml[1]).toContain(
      '<dc:title>Quarterly &amp; &lt;Review&gt;</dc:title>',
    );
    expect(coreXml[2]).toContain('<dc:title></dc:title>');

    const reopened = await Promise.all(imported.map(async (document) =>
      PptxDocument.open(await document.write())));
    expect(reopened.map(({ title }) => title)).toEqual(expectedTitles);

    const native = PptxDocument.create({ title: 'Quarterly & <Review>' });
    const nativeOmitted = PptxDocument.create();
    expect(native.title).toBe(imported[1]!.title);
    expect(nativeOmitted.title).toBeUndefined();
    expect(new TextDecoder().decode(
      native.opcPackage.requirePart('/docProps/core.xml').bytes,
    )).toContain('<dc:title>Quarterly &amp; &lt;Review&gt;</dc:title>');
    const reopenedNative = await PptxDocument.open(await native.write());
    expect(reopenedNative.title).toBe('Quarterly & <Review>');
  }, 20_000);

  it('imports and reopens PptxGenJS presentation author metadata from public output', async () => {
    const baseline = new PptxGenJS();
    baseline.addSlide();
    const custom = new PptxGenJS();
    custom.author = 'Alice & <Bob>';
    custom.addSlide();
    const empty = new PptxGenJS();
    empty.author = '';
    empty.addSlide();
    expect([baseline.version, custom.version, empty.version]).toEqual([
      '4.0.1',
      '4.0.1',
      '4.0.1',
    ]);

    const expectedAuthors = [
      'PptxGenJS',
      'Alice & <Bob>',
      '',
    ] as const;
    const imported = await Promise.all([
      importPptxGenJS(baseline),
      importPptxGenJS(custom),
      importPptxGenJS(empty),
    ]);
    expect(imported.map(({ author }) => author)).toEqual(expectedAuthors);
    for (const [index, document] of imported.entries()) {
      const journal = [...document.opcPackage.mutations];
      expect(document.author).toBe(expectedAuthors[index]);
      expect(document.opcPackage.mutations).toEqual(journal);
    }

    const coreXml = imported.map((document) => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/core.xml').bytes,
    ));
    expect(coreXml[0]).toContain('<dc:creator>PptxGenJS</dc:creator>');
    expect(coreXml[0]).toContain('<cp:lastModifiedBy>PptxGenJS</cp:lastModifiedBy>');
    expect(coreXml[1]).toContain(
      '<dc:creator>Alice &amp; &lt;Bob&gt;</dc:creator>',
    );
    expect(coreXml[1]).toContain(
      '<cp:lastModifiedBy>Alice &amp; &lt;Bob&gt;</cp:lastModifiedBy>',
    );
    expect(coreXml[2]).toContain('<dc:creator></dc:creator>');
    expect(coreXml[2]).toContain('<cp:lastModifiedBy></cp:lastModifiedBy>');

    const reopened = await Promise.all(imported.map(async (document) =>
      PptxDocument.open(await document.write())));
    expect(reopened.map(({ author }) => author)).toEqual(expectedAuthors);

    const native = PptxDocument.create({ author: 'Alice & <Bob>' });
    const nativeOmitted = PptxDocument.create();
    expect(native.author).toBe(imported[1]!.author);
    expect(nativeOmitted.author).toBe('@jiayunxie/pptx');
    const nativeCore = new TextDecoder().decode(
      native.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    expect(nativeCore).toContain(
      '<dc:creator>Alice &amp; &lt;Bob&gt;</dc:creator>',
    );
    expect(nativeCore).toContain(
      '<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>',
    );
    const reopenedNative = await PptxDocument.open(await native.write());
    expect(reopenedNative.author).toBe('Alice & <Bob>');
    expect(new TextDecoder().decode(
      reopenedNative.opcPackage.requirePart('/docProps/core.xml').bytes,
    )).toContain('<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>');
  }, 20_000);

  it('imports and reopens PptxGenJS presentation last modified by metadata from public output', async () => {
    const baseline = new PptxGenJS();
    baseline.addSlide();
    const custom = new PptxGenJS();
    custom.author = 'Alice & <Bob>';
    custom.addSlide();
    const empty = new PptxGenJS();
    empty.author = '';
    empty.addSlide();
    expect([baseline.version, custom.version, empty.version]).toEqual([
      '4.0.1',
      '4.0.1',
      '4.0.1',
    ]);

    const expectedEditors = ['PptxGenJS', 'Alice & <Bob>', ''] as const;
    const imported = await Promise.all([
      importPptxGenJS(baseline),
      importPptxGenJS(custom),
      importPptxGenJS(empty),
    ]);
    expect(imported.map(({ lastModifiedBy }) => lastModifiedBy)).toEqual(expectedEditors);
    expect(imported.map(({ author }) => author)).toEqual(expectedEditors);
    for (const [index, document] of imported.entries()) {
      const journal = [...document.opcPackage.mutations];
      expect(document.lastModifiedBy).toBe(expectedEditors[index]);
      expect(document.author).toBe(expectedEditors[index]);
      expect(document.opcPackage.mutations).toEqual(journal);
    }

    const coreXml = imported.map((document) => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/core.xml').bytes,
    ));
    expect(coreXml[0]).toContain('<dc:creator>PptxGenJS</dc:creator>');
    expect(coreXml[0]).toContain('<cp:lastModifiedBy>PptxGenJS</cp:lastModifiedBy>');
    expect(coreXml[1]).toContain(
      '<dc:creator>Alice &amp; &lt;Bob&gt;</dc:creator>',
    );
    expect(coreXml[1]).toContain(
      '<cp:lastModifiedBy>Alice &amp; &lt;Bob&gt;</cp:lastModifiedBy>',
    );
    expect(coreXml[2]).toContain('<dc:creator></dc:creator>');
    expect(coreXml[2]).toContain('<cp:lastModifiedBy></cp:lastModifiedBy>');

    const reopened = await Promise.all(imported.map(async (document) =>
      PptxDocument.open(await document.write())));
    expect(reopened.map(({ lastModifiedBy }) => lastModifiedBy)).toEqual(expectedEditors);
    expect(reopened.map(({ author }) => author)).toEqual(expectedEditors);

    const nativeDefault = PptxDocument.create();
    expect(nativeDefault.author).toBe('@jiayunxie/pptx');
    expect(nativeDefault.lastModifiedBy).toBe('@jiayunxie/pptx');
    expect(nativeDefault.lastModifiedBy).not.toBe(imported[0]!.lastModifiedBy);

    const nativeMirror = PptxDocument.create({
      author: 'Alice & <Bob>',
      lastModifiedBy: 'Alice & <Bob>',
    });
    expect(nativeMirror.author).toBe(imported[1]!.author);
    expect(nativeMirror.lastModifiedBy).toBe(imported[1]!.lastModifiedBy);
    const nativeCore = new TextDecoder().decode(
      nativeMirror.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    expect(nativeCore).toContain(
      '<dc:creator>Alice &amp; &lt;Bob&gt;</dc:creator>',
    );
    expect(nativeCore).toContain(
      '<cp:lastModifiedBy>Alice &amp; &lt;Bob&gt;</cp:lastModifiedBy>',
    );

    nativeMirror.author = 'Creator only';
    expect(nativeMirror.author).toBe('Creator only');
    expect(nativeMirror.lastModifiedBy).toBe('Alice & <Bob>');
    nativeMirror.lastModifiedBy = 'Editor only';
    expect(nativeMirror.author).toBe('Creator only');
    expect(nativeMirror.lastModifiedBy).toBe('Editor only');
    const reopenedNative = await PptxDocument.open(await nativeMirror.write());
    expect(reopenedNative.author).toBe('Creator only');
    expect(reopenedNative.lastModifiedBy).toBe('Editor only');
  }, 20_000);

  it('imports and reopens PptxGenJS presentation subject metadata from public output', async () => {
    const baseline = new PptxGenJS();
    baseline.addSlide();
    const custom = new PptxGenJS();
    custom.subject = 'Revenue & <Forecast>';
    custom.addSlide();
    const empty = new PptxGenJS();
    empty.subject = '';
    empty.addSlide();
    expect([baseline.version, custom.version, empty.version]).toEqual([
      '4.0.1',
      '4.0.1',
      '4.0.1',
    ]);

    const expectedSubjects = [
      'PptxGenJS Presentation',
      'Revenue & <Forecast>',
      '',
    ] as const;
    const imported = await Promise.all([
      importPptxGenJS(baseline),
      importPptxGenJS(custom),
      importPptxGenJS(empty),
    ]);
    expect(imported.map(({ subject }) => subject)).toEqual(expectedSubjects);
    for (const [index, document] of imported.entries()) {
      const journal = [...document.opcPackage.mutations];
      expect(document.subject).toBe(expectedSubjects[index]);
      expect(document.opcPackage.mutations).toEqual(journal);
    }

    const coreXml = imported.map((document) => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/core.xml').bytes,
    ));
    expect(coreXml[0]).toContain(
      '<dc:subject>PptxGenJS Presentation</dc:subject>',
    );
    expect(coreXml[1]).toContain(
      '<dc:subject>Revenue &amp; &lt;Forecast&gt;</dc:subject>',
    );
    expect(coreXml[2]).toContain('<dc:subject></dc:subject>');

    const reopened = await Promise.all(imported.map(async (document) =>
      PptxDocument.open(await document.write())));
    expect(reopened.map(({ subject }) => subject)).toEqual(expectedSubjects);

    const native = PptxDocument.create({ subject: 'Revenue & <Forecast>' });
    const nativeOmitted = PptxDocument.create();
    expect(native.subject).toBe(imported[1]!.subject);
    expect(nativeOmitted.subject).toBeUndefined();
    expect(new TextDecoder().decode(
      nativeOmitted.opcPackage.requirePart('/docProps/core.xml').bytes,
    )).not.toContain('<dc:subject');
    expect(new TextDecoder().decode(
      native.opcPackage.requirePart('/docProps/core.xml').bytes,
    )).toContain('<dc:subject>Revenue &amp; &lt;Forecast&gt;</dc:subject>');
    const reopenedNative = await PptxDocument.open(await native.write());
    expect(reopenedNative.subject).toBe('Revenue & <Forecast>');
    expect(new TextDecoder().decode(
      reopenedNative.opcPackage.requirePart('/docProps/core.xml').bytes,
    )).toContain('<dc:subject>Revenue &amp; &lt;Forecast&gt;</dc:subject>');
  }, 20_000);

  it('imports and reopens PptxGenJS presentation revision metadata from public output', async () => {
    const baseline = new PptxGenJS();
    baseline.addSlide();
    const zero = new PptxGenJS();
    zero.revision = '0';
    zero.addSlide();
    const custom = new PptxGenJS();
    custom.revision = '42';
    custom.addSlide();
    const leading = new PptxGenJS();
    leading.revision = '007';
    leading.addSlide();
    expect([baseline.version, zero.version, custom.version, leading.version]).toEqual([
      '4.0.1',
      '4.0.1',
      '4.0.1',
      '4.0.1',
    ]);

    const expectedRevisions = ['1', '0', '42', '007'] as const;
    const imported = await Promise.all([
      importPptxGenJS(baseline),
      importPptxGenJS(zero),
      importPptxGenJS(custom),
      importPptxGenJS(leading),
    ]);
    expect(imported.map(({ revision }) => revision)).toEqual(expectedRevisions);
    for (const [index, document] of imported.entries()) {
      const journal = [...document.opcPackage.mutations];
      expect(document.revision).toBe(expectedRevisions[index]);
      expect(document.opcPackage.mutations).toEqual(journal);
    }

    const coreXml = imported.map((document) => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/core.xml').bytes,
    ));
    expect(coreXml[0]).toContain('<cp:revision>1</cp:revision>');
    expect(coreXml[1]).toContain('<cp:revision>0</cp:revision>');
    expect(coreXml[2]).toContain('<cp:revision>42</cp:revision>');
    expect(coreXml[3]).toContain('<cp:revision>007</cp:revision>');

    const reopened = await Promise.all(imported.map(async (document) =>
      PptxDocument.open(await document.write())));
    expect(reopened.map(({ revision }) => revision)).toEqual(expectedRevisions);

    const native = PptxDocument.create({ revision: '42' });
    const nativeOmitted = PptxDocument.create();
    expect(native.revision).toBe(imported[2]!.revision);
    expect(nativeOmitted.revision).toBe(imported[0]!.revision);
    expect(new TextDecoder().decode(
      native.opcPackage.requirePart('/docProps/core.xml').bytes,
    )).toContain('<cp:revision>42</cp:revision>');
    const reopenedNative = await PptxDocument.open(await native.write());
    expect(reopenedNative.revision).toBe('42');

    const invalidValues = ['', '-1', '1.5', 'abc'] as const;
    for (const value of invalidValues) {
      const generated = new PptxGenJS();
      generated.revision = value;
      generated.addSlide();
      const document = await importPptxGenJS(generated);
      const before = document.opcPackage.requirePart('/docProps/core.xml').bytes;
      expect(document.revision).toBeUndefined();
      expect(new TextDecoder().decode(before)).toContain(
        `<cp:revision>${value}</cp:revision>`,
      );
      const invalidReopened = await PptxDocument.open(await document.write());
      expect(invalidReopened.revision).toBeUndefined();
      expect(invalidReopened.opcPackage.requirePart('/docProps/core.xml').bytes)
        .toEqual(before);
      expect(() => PptxDocument.create({ revision: value })).toThrow(TypeError);
    }
  }, 20_000);

  it('imports and reopens PptxGenJS presentation company metadata from public output', async () => {
    const baseline = new PptxGenJS();
    baseline.addSlide();
    const custom = new PptxGenJS();
    custom.company = 'Acme 国际';
    custom.addSlide();
    const empty = new PptxGenJS();
    empty.company = '';
    empty.addSlide();
    expect([baseline.version, custom.version, empty.version]).toEqual([
      '4.0.1',
      '4.0.1',
      '4.0.1',
    ]);

    const expectedCompanies = [
      'PptxGenJS',
      'Acme 国际',
      '',
    ] as const;
    const imported = await Promise.all([
      importPptxGenJS(baseline),
      importPptxGenJS(custom),
      importPptxGenJS(empty),
    ]);
    expect(imported.map(({ company }) => company)).toEqual(expectedCompanies);
    for (const [index, document] of imported.entries()) {
      const journal = [...document.opcPackage.mutations];
      expect(document.company).toBe(expectedCompanies[index]);
      expect(document.opcPackage.mutations).toEqual(journal);
    }

    const appXml = imported.map((document) => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/app.xml').bytes,
    ));
    expect(appXml[0]).toContain('<Company>PptxGenJS</Company>');
    expect(appXml[1]).toContain('<Company>Acme 国际</Company>');
    expect(appXml[2]).toContain('<Company></Company>');
    for (const xml of appXml) {
      expect(xml).toContain('<Application>Microsoft Office PowerPoint</Application>');
      expect(xml).toContain('<AppVersion>16.0000</AppVersion>');
    }

    const reopened = await Promise.all(imported.map(async (document) =>
      PptxDocument.open(await document.write())));
    expect(reopened.map(({ company }) => company)).toEqual(expectedCompanies);

    const native = PptxDocument.create({ company: 'Acme & <Partners>' });
    const nativeOmitted = PptxDocument.create();
    expect(native.company).toBe('Acme & <Partners>');
    expect(nativeOmitted.company).toBeUndefined();
    const nativeApp = new TextDecoder().decode(
      native.opcPackage.requirePart('/docProps/app.xml').bytes,
    );
    expect(nativeApp).toContain(
      '<Company>Acme &amp; &lt;Partners&gt;</Company>',
    );
    expect(nativeApp).toContain('<Application>@jiayunxie/pptx</Application>');
    expect(nativeApp).toContain('<AppVersion>1.0</AppVersion>');
    const reopenedNative = await PptxDocument.open(await native.write());
    expect(reopenedNative.company).toBe('Acme & <Partners>');

    const unsafe = new PptxGenJS();
    unsafe.company = 'A & <B>';
    unsafe.addSlide();
    const importedUnsafe = await importPptxGenJS(unsafe);
    const unsafeApp = new TextDecoder().decode(
      importedUnsafe.opcPackage.requirePart('/docProps/app.xml').bytes,
    );
    expect(unsafeApp).toContain('<Company>A & <B></Company>');
    expect(importedUnsafe.company).toBeUndefined();
  }, 20_000);

  it('imports PptxGenJS non-list zero margins and indents without aliasing bullet indentation', async () => {
    const generated = new PptxGenJS();
    const slide = generated.addSlide();
    slide.addText('Plain', { name: 'Margin plain', x: 1, y: 1, w: 3, h: 0.5 });
    slide.addText([{ text: 'Rich' }], { name: 'Margin rich', x: 1, y: 2, w: 3, h: 0.5 });
    slide.addText('Bullet', {
      name: 'Margin bullet',
      x: 1,
      y: 3,
      w: 3,
      h: 0.5,
      bullet: true,
    });
    slide.addText('Number', {
      name: 'Margin number',
      x: 1,
      y: 4,
      w: 3,
      h: 0.5,
      bullet: { type: 'number', numberType: 'romanUcPeriod', numberStartAt: 1, indent: 22 },
    });
    const document = await importPptxGenJS(generated);
    const shapes = document.slides[0]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel);

    expect(shapes.map(({ richText }) => richText[0]?.marginLeft)).toEqual([
      0,
      0,
      undefined,
      undefined,
    ]);
    expect(shapes.map(({ richText }) => richText[0]?.marginRight)).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    expect(shapes.map(({ richText }) => richText[0]?.indent)).toEqual([
      0,
      0,
      undefined,
      undefined,
    ]);
    expect(shapes[2]!.richText[0]!.bullet).toEqual({ kind: 'bullet', character: '•', indent: 27 });
    expect(shapes[3]!.richText[0]!.bullet).toEqual({
      kind: 'number',
      style: 'arabicPeriod',
      startAt: 1,
      indent: 22,
    });
    const slideXml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(slideXml.match(/indent="0" marL="0"/g)).toHaveLength(2);
    expect(slideXml).toMatch(/marL="342900" indent="-342900"/);
    expect(slideXml).toMatch(/marL="279400" indent="-279400"/);
    expect(slideXml).not.toContain('marR=');

    const reopened = await PptxDocument.open(await document.write());
    const reopenedShapes = reopened.slides[0]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel);
    expect(reopenedShapes.map(({ richText }) => richText[0]?.marginLeft)).toEqual([
      0,
      0,
      undefined,
      undefined,
    ]);
    expect(reopenedShapes.map(({ richText }) => richText[0]?.marginRight)).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    expect(reopenedShapes.map(({ richText }) => richText[0]?.indent)).toEqual([
      0,
      0,
      undefined,
      undefined,
    ]);
  }, 20_000);

  it('keeps pptxgenjs out of every non-adapter package dependency list', async () => {
    const packagesDirectory = fileURLToPath(new URL('../..', import.meta.url));
    const packageNames = ['lossless-xml', 'model', 'opc', 'sdk', 'validator'];
    for (const packageName of packageNames) {
      const manifest = JSON.parse(await readFile(`${packagesDirectory}/${packageName}/package.json`, 'utf8')) as {
        dependencies?: Record<string, string>;
      };
      expect(manifest.dependencies?.pptxgenjs, packageName).toBeUndefined();
    }
  });
});
