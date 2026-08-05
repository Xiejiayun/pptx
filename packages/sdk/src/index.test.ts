import { createHash } from 'node:crypto';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Readable } from 'node:stream';
import JSZip from 'jszip';
import { describe, expect, it, vi } from 'vitest';
import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import {
  PRESENTATION_FORMAT_PROFILES,
  type NumberingStyle,
  type PresentationFormat,
  type RichTextStrikeStyle,
  type RichTextUnderlineStyle,
} from '@pptx/model';
import { OpcPackage } from '@pptx/opc';
import { validatePackage } from '@pptx/validator';
import {
  ChartModel,
  CustomGeometryEvaluationError,
  degrees,
  evaluateCustomGeometry,
  ImageModel,
  inches,
  inspectRasterImage,
  MediaModel,
  ModelParseError,
  openPptxStream,
  OUTPUT_TYPES,
  PRESET_SHAPE_TYPES,
  TEXT_ALIGNMENTS,
  TEXT_VERTICAL_ALIGNMENTS,
  PPTX_VERSION,
  PptxDocument,
  ShapeModel,
  SlideLayoutModel,
  SlideMasterModel,
  TableModel,
  ValidationError,
  type AddImageOptions,
  type AddChartOptions,
  type AddImageSourceOptions,
  type AddMediaOptions,
  type AddSvgImageOptions,
  type AddTableCell,
  type AddTableCellOptions,
  type AddTableCellInput,
  type AddTableOptions,
  type CustomGeometry,
  type ChartType,
  type CustomGeometryEvaluationContext,
  type EvaluatedCustomGeometry,
  type Hyperlink,
  type ImageSource,
  type ImageSizing,
  type InsertTableColumnsOptions,
  type InsertTableRowsOptions,
  type MediaKind,
  type MediaPlaybackSettings,
  type MediaSource,
  type OutputType,
  type RasterImageContentType,
  type RasterImageByteStream,
  type RasterImageInfo,
  type RasterImageSource,
  type RichTextParagraph,
  type SetSlideBackgroundImageOptions,
  type SlideModel,
  type SvgImageContentType,
  type ShapeArrows,
  type ShapeAdjustment,
  type ShapeArrowType,
  type ShapeFill,
  type ShapeLine,
  type ShapeLineDash,
  type ShapeShadow,
  type CreatePresentationOptions,
  type SlideNumber,
  type SlideNumberColor,
  type SlideNumberMarginInput,
  type SlideNumberMargins,
  type SlideNumberOptions,
  type SlideNumberTextStyle,
  type SlideNumberTextStyleOptions,
  type TextAlignment,
  type TextBoxMargins,
  type TableCellBorders,
  type TableCell,
  type TableCellMerge,
  type TableCellFill,
  type TableCellTextDirection,
  type TableMergeRegion,
  type TableAutoPageMarginInput,
  type TableToSlidesAddImage,
  type TableToSlidesAddShape,
  type TableToSlidesAddTable,
  type TableToSlidesAddText,
  type TableToSlidesOptions,
  type TextBoxVerticalAlignment,
  type PresentationLayout,
  type PresentationLayoutName,
  type PptxVersion,
  type WriteBaseOptions,
  type WriteOptions,
  type WriteOutput,
} from './index.js';

function sdkPngHeader(width: number, height: number): Uint8Array<ArrayBuffer> {
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

function sdkSvg(width = 640, height = 360): Uint8Array<ArrayBuffer> {
  return new TextEncoder().encode(
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${width} ${height}"/>`,
  );
}

function sdkPngDataUri(bytes: Uint8Array): string {
  return `data:image/png;base64,${Buffer.from(bytes).toString('base64')}`;
}

function sdkGif(width: number, height: number): Uint8Array<ArrayBuffer> {
  return Uint8Array.from([
    71, 73, 70, 56, 57, 97,
    width & 0xff,
    width >>> 8,
    height & 0xff,
    height >>> 8,
  ]);
}

function sdkJpeg(width: number, height: number): Uint8Array<ArrayBuffer> {
  return Uint8Array.from([
    0xff, 0xd8,
    0xff, 0xc0, 0x00, 0x08, 0x08,
    height >>> 8, height & 0xff,
    width >>> 8, width & 0xff,
    0x01,
  ]);
}

function sdkSlideNumberCache(document: PptxDocument, ownerPartUri: string): string | undefined {
  const xml = LosslessXmlDocument.parse(document.opcPackage.requirePart(ownerPartUri).bytes);
  const field = xml.elements('fld').find(
    (candidate) => xml.attribute(candidate, 'type')?.value === 'slidenum',
  );
  const texts = field ? xml.descendants(field, 't') : [];
  return texts.length === 1 ? xml.text(texts[0]!) : undefined;
}

async function titleFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId7"/></p:sldIdLst></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Original</a:t></a:r></a:p></p:txBody><x:unknown xmlns:x="x" custom="keep"/></p:sp></p:spTree></p:cSld></p:sld>');
  zip.file('ppt/theme/theme1.xml', '<a:theme xmlns:a="a"><x:opaque xmlns:x="x">KEEP</x:opaque></a:theme>');
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

async function tableTextDirectionFixture(): Promise<Uint8Array> {
  const document = PptxDocument.create();
  const slide = document.addSlide();
  const part = document.opcPackage.requirePart(slide.partUri);
  const cell = (
    text: string,
    properties: string,
    attributes = '',
    bodyProperties = '<a:bodyPr/>',
  ): string =>
    `<a:tc${attributes}><a:txBody>${bodyProperties}<a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></a:txBody>${properties}</a:tc>`;
  const table = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Direction table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="7315200" cy="2743200"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid><a:gridCol w="1463040"/><a:gridCol w="1463040"/><a:gridCol w="1463040"/><a:gridCol w="1463040"/><a:gridCol w="1463040"/></a:tblGrid>'
    + `<a:tr h="1371600">${[
      cell('Horz', '<a:tcPr vert="horz"/>'),
      cell('Vert', '<a:tcPr vert="vert"/>'),
      cell('Vert 270', '<a:tcPr vert="vert270"/>'),
      cell('WordArt', '<a:tcPr vert="wordArtVert"/>'),
      cell('Absent', '<a:tcPr keep="ABSENT"/>'),
    ].join('')}</a:tr>`
    + `<a:tr h="1371600">${[
      cell('Clear me', '<a:tcPr vert="vert"/>'),
      cell('Edit me', '<a:tcPr marL="100"><x:keep xmlns:x="urn:test">OPAQUE</x:keep></a:tcPr>'),
      cell('Merged placeholder', '<a:tcPr marR="200"/>', ' hMerge="1"'),
      cell('Neighbor', '<a:tcPr vert="horz" keep="NEIGHBOR"/>'),
      cell('Tail', '<a:tcPr/>'),
    ].join('')}</a:tr>`
    + `<a:tr h="1371600">${[
      cell('Explicit none', '<a:tcPr vert="horz" anchor="t" marT="50800" marR="101600" marB="152400" marL="203200" keep="FIT-NONE"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:tcPr>', '', '<a:bodyPr custom="NONE"><a:noAutofit/></a:bodyPr>'),
      cell('Calculated shrink', '<a:tcPr vert="vert" anchor="ctr" marT="0" marL="25400" keep="FIT-SHRINK"><a:solidFill><a:schemeClr val="accent2"><a:alpha val="75000"/></a:schemeClr></a:solidFill></a:tcPr>', '', '<a:bodyPr custom="SHRINK"><a:normAutofit fontScale="85000" lnSpcReduction="20000"/></a:bodyPr>'),
      cell('Resize', '<a:tcPr vert="vert270" anchor="b" marR="91440" keep="FIT-RESIZE"><a:noFill/></a:tcPr>', '', '<a:bodyPr custom="RESIZE"><a:spAutoFit/></a:bodyPr>'),
      cell('Absent fit', '<a:tcPr keep="FIT-ABSENT"/>'),
      cell('Merged fit', '<a:tcPr vert="wordArtVert" anchor="ctr" marT="45720" marR="91440" marB="45720" marL="91440" keep="FIT-MERGED"><a:solidFill><a:srgbClr val="70AD47"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:tcPr>', ' hMerge="1"'),
    ].join('')}</a:tr>`
    + '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
  document.opcPackage.setPart(
    slide.partUri,
    new TextDecoder().decode(part.bytes).replace('</p:spTree>', `${table}</p:spTree>`),
    part.contentType,
  );
  return document.opcPackage.write();
}

async function tableHorizontalAlignmentFixture(): Promise<Uint8Array> {
  const document = await PptxDocument.open(await tableTextDirectionFixture());
  const slide = document.slides[0]!;
  const part = document.opcPackage.requirePart(slide.partUri);
  let source = new TextDecoder().decode(part.bytes);
  for (const [text, properties] of [
    ['Explicit none', '<a:pPr algn="l" keep="LEFT"/>'],
    ['Calculated shrink', '<a:pPr algn="ctr" keep="CENTER"><a:buNone/></a:pPr>'],
    ['Resize', '<a:pPr algn="r" keep="RIGHT"/>'],
    ['Absent fit', '<a:pPr algn="just" keep="JUSTIFY"/>'],
    ['Merged fit', '<a:pPr keep="ABSENT"/>'],
  ] as const) {
    const target = `<a:p><a:r><a:rPr lang="en-US"/><a:t>${text}</a:t>`;
    expect(source.split(target)).toHaveLength(2);
    source = source.replace(
      target,
      `<a:p>${properties}<a:r><a:rPr lang="en-US"/><a:t>${text}</a:t>`,
    );
  }
  document.opcPackage.setPart(slide.partUri, source, part.contentType);
  return document.write();
}

async function tableBordersFixture(): Promise<Uint8Array> {
  const document = await PptxDocument.open(await tableTextDirectionFixture());
  const slide = document.slides[0]!;
  const part = document.opcPackage.requirePart(slide.partUri);
  const line = (
    tag: 'lnL' | 'lnR' | 'lnT' | 'lnB',
    width: number,
    fill: string,
    dash = '',
  ): string =>
    `<a:${tag} w="${width}" cap="flat" cmpd="sng" algn="ctr">${fill}${dash}<a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:${tag}>`;
  const red = ['lnL', 'lnR', 'lnT', 'lnB'].map((tag) =>
    line(
      tag as 'lnL' | 'lnR' | 'lnT' | 'lnB',
      12700,
      '<a:solidFill><a:srgbClr val="C00000"/></a:solidFill>',
      '<a:prstDash val="solid"/>',
    )).join('');
  const none = (tag: 'lnL' | 'lnR' | 'lnT' | 'lnB'): string =>
    `<a:${tag} w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:${tag}>`;
  const fourNone = [none('lnL'), none('lnR'), none('lnT'), none('lnB')].join('');
  const mixed = line(
    'lnR',
    0,
    '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>',
  ) + line(
    'lnT',
    19050,
    '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>',
    '<a:prstDash val="sysDash"/>',
  ) + none('lnB');
  const merged = line(
    'lnL',
    25400,
    '<a:solidFill><a:srgbClr val="333333"/></a:solidFill>',
    '<a:prstDash val="solid"/>',
  );
  const source = new TextDecoder().decode(part.bytes)
    .replace('keep="FIT-NONE"><a:solidFill>', `keep="FIT-NONE">${red}<a:solidFill>`)
    .replace('keep="FIT-SHRINK"><a:solidFill>', `keep="FIT-SHRINK">${mixed}<a:solidFill>`)
    .replace('keep="FIT-RESIZE"><a:noFill/>', `keep="FIT-RESIZE">${fourNone}<a:lnTlToBr w="12700"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:lnTlToBr><a:noFill/>`)
    .replace('keep="FIT-MERGED"><a:solidFill>', `keep="FIT-MERGED">${merged}<a:solidFill>`);
  document.opcPackage.setPart(slide.partUri, source, part.contentType);
  return document.write();
}

describe('PptxDocument vertical slice', () => {
  it('exposes a read-only runtime version without package mutation', async () => {
    const current: PptxVersion = PPTX_VERSION;
    const document = PptxDocument.create();
    const before = await sdkPackageSnapshot(document);

    expect(current).toBe('0.1.0');
    expect(document.version).toBe(current);
    expect(document.version).toBe(document.version);
    expect(typeof PptxDocument.prototype.tableToSlides).toBe('function');
    expect(await sdkPackageSnapshot(document)).toEqual(before);
    expect(Object.getOwnPropertyDescriptor(PptxDocument.prototype, 'version')).toMatchObject({
      set: undefined,
      enumerable: false,
    });

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.version).toBe(PPTX_VERSION);

    if (false) {
      const addImage: TableToSlidesAddImage = { source: sdkPngHeader(1, 1) };
      const addShape: TableToSlidesAddShape = { type: 'rect' };
      const addTable: TableToSlidesAddTable = { rows: [['A']] };
      const addText: TableToSlidesAddText = {
        text: [{ runs: [{ text: 'Title', style: { bold: true } }] }],
      };
      const options: TableToSlidesOptions = {
        autoPage: true,
        addImage,
        addShape,
        addTable,
        addText,
      };
      const pages: Promise<readonly SlideModel[]> = document.tableToSlides('table', options);
      // @ts-expect-error version is read-only
      document.version = '9.9.9';
      void pages;
    }
  });

  it('publishes the frozen output type catalog without package mutation', () => {
    const document = PptxDocument.create();
    const journal = [...document.opcPackage.mutations];
    const values: readonly OutputType[] = OUTPUT_TYPES;

    expect(values).toBe(OUTPUT_TYPES);
    expect([...values]).toEqual([
      'arraybuffer',
      'base64',
      'binarystring',
      'blob',
      'nodebuffer',
      'uint8array',
    ]);
    expect(Object.isFrozen(OUTPUT_TYPES)).toBe(true);
    expect(document.opcPackage.mutations).toEqual(journal);
  });

  it('writes every selected output representation without changing canonical bytes', async () => {
    const document = PptxDocument.create();
    document.addSlide().addText('Output types 你好');
    const journal = [...document.opcPackage.mutations];
    const defaultOutput = await document.write();
    const emptyOutput = await document.write({});
    const permissiveOutput = await document.write({ mode: 'permissive' });
    const explicitUndefined = await document.write({
      outputType: undefined,
    } as unknown as WriteOptions);
    const arraybuffer = await document.write({ outputType: 'arraybuffer' });
    const base64 = await document.write({ outputType: 'base64' });
    const binarystring = await document.write({ outputType: 'binarystring' });
    const blob = await document.write({ outputType: 'blob' });
    const nodebuffer = await document.write({ outputType: 'nodebuffer' });
    const uint8array = await document.write({ outputType: 'uint8array' });

    expect(Buffer.isBuffer(defaultOutput)).toBe(false);
    expect(emptyOutput).toEqual(defaultOutput);
    expect(permissiveOutput).toEqual(defaultOutput);
    expect(explicitUndefined).toEqual(defaultOutput);
    expect(new Uint8Array(arraybuffer)).toEqual(defaultOutput);
    expect(Uint8Array.from(Buffer.from(base64, 'base64'))).toEqual(defaultOutput);
    expect(Uint8Array.from(binarystring, (value) => value.charCodeAt(0))).toEqual(defaultOutput);
    expect(new Uint8Array(await blob.arrayBuffer())).toEqual(defaultOutput);
    expect(blob.type).toBe('application/zip');
    expect(Buffer.isBuffer(nodebuffer)).toBe(true);
    expect(new Uint8Array(nodebuffer)).toEqual(defaultOutput);
    expect(Buffer.isBuffer(uint8array)).toBe(false);
    expect(uint8array).toEqual(defaultOutput);
    expect(document.opcPackage.mutations).toEqual(journal);

    const decoded = [
      new Uint8Array(arraybuffer),
      Uint8Array.from(Buffer.from(base64, 'base64')),
      Uint8Array.from(binarystring, (value) => value.charCodeAt(0)),
      new Uint8Array(await blob.arrayBuffer()),
      nodebuffer,
      uint8array,
    ];
    for (const bytes of decoded) {
      const reopened = await PptxDocument.open(bytes);
      const shape = reopened.slides[0]?.shapes[0];
      expect(shape).toBeInstanceOf(ShapeModel);
      expect((shape as ShapeModel).text).toBe('Output types 你好');
    }

    const convenienceBlob = await document.writeBlob();
    expect(convenienceBlob.type).toBe(
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    );
    expect(new Uint8Array(await convenienceBlob.arrayBuffer())).toEqual(defaultOutput);

    if (false) {
      const baseOptions: WriteBaseOptions = { mode: 'permissive' };
      const blobOptions: WriteOptions<'blob'> = { outputType: 'blob' };
      const dynamicOptions: WriteOptions<OutputType> = { outputType: OUTPUT_TYPES[0] };
      document.write() satisfies Promise<Uint8Array>;
      document.write(baseOptions) satisfies Promise<Uint8Array>;
      document.write(blobOptions) satisfies Promise<Blob>;
      document.write(dynamicOptions) satisfies Promise<WriteOutput<OutputType>>;
      document.write({ outputType: 'arraybuffer' }) satisfies Promise<ArrayBuffer>;
      document.write({ outputType: 'base64' }) satisfies Promise<string>;
      document.write({ outputType: 'binarystring' }) satisfies Promise<string>;
      document.write({ outputType: 'nodebuffer' }) satisfies Promise<Uint8Array>;
      document.write({ outputType: 'uint8array' }) satisfies Promise<Uint8Array>;
      document.writeBlob(baseOptions) satisfies Promise<Blob>;
      // @ts-expect-error convenience blob output has a fixed presentation MIME contract
      document.writeBlob({ outputType: 'blob' });
      // @ts-expect-error generic blob options accept only the blob token
      const wrongBlobOptions: WriteOptions<'blob'> = { outputType: 'base64' };
      void wrongBlobOptions;
    }
  });

  it('rejects unsupported write output types before diagnostics or package writes', async () => {
    for (const outputType of ['STREAM', 'buffer', 'BLOB', null, 1, {}]) {
      const document = PptxDocument.create();
      const diagnostics = [...document.diagnostics];
      const journal = [...document.opcPackage.mutations];
      const write = vi.spyOn(document.opcPackage, 'write');

      await expect(document.write({ outputType } as never)).rejects.toThrow(
        new TypeError('PptxDocument.write() received an unsupported outputType'),
      );
      expect(write).not.toHaveBeenCalled();
      expect(document.diagnostics).toEqual(diagnostics);
      expect(document.opcPackage.mutations).toEqual(journal);
      write.mockRestore();
    }
  });

  it('exposes a detached presentation layout projection across its lifecycle', async () => {
    const sizes = [
      ['4:3', 'screen4x3', 9_144_000, 6_858_000],
      ['16:9', 'screen16x9', 9_144_000, 5_143_500],
      ['16:10', 'screen16x10', 9_144_000, 5_715_000],
      ['wide', 'custom', 12_192_000, 6_858_000],
    ] as const;

    for (const [slideSize, name, width, height] of sizes) {
      const document = PptxDocument.create({ slideSize });
      const before = await sdkPackageSnapshot(document);
      const layout: PresentationLayout = document.presLayout;
      const layoutName: PresentationLayoutName = layout.name;
      expect({ ...layout, name: layoutName }).toEqual({ name, width, height });
      expect(document.presLayout).not.toBe(layout);
      expect(await sdkPackageSnapshot(document)).toEqual(before);
    }

    const custom = PptxDocument.create({
      slideSize: { width: inches(11.7), height: inches(8.3) },
    });
    const detached = custom.presLayout as {
      name: string;
      width: number;
      height: number;
    };
    detached.name = 'changed';
    detached.width = 1;
    expect(custom.presLayout).toEqual({
      name: 'custom',
      width: inches(11.7),
      height: inches(8.3),
    });

    custom.slideSize = { width: inches(10), height: inches(6.25) };
    const edited: PresentationLayout = custom.presLayout;
    expect(edited).toEqual({
      name: 'screen16x10',
      width: inches(10),
      height: inches(6.25),
    });
    expect(() => {
      custom.slideSize = { width: 0 as never, height: inches(7.5) };
    }).toThrow(RangeError);
    expect(custom.presLayout).toEqual(edited);
    expect((await PptxDocument.open(await custom.write())).presLayout).toEqual(edited);
    expect(Object.getOwnPropertyDescriptor(PptxDocument.prototype, 'presLayout')).toMatchObject({
      set: undefined,
      enumerable: false,
    });

    const malformed = PptxDocument.create();
    const presentationPart = malformed.opcPackage.requirePart(malformed.presentationPartUri);
    malformed.opcPackage.setPart(
      presentationPart.uri,
      new TextDecoder().decode(presentationPart.bytes).replace('cx="9144000"', 'cx="0"'),
      presentationPart.contentType,
    );
    expect(() => malformed.presLayout).toThrow(/slide width is invalid/);

    if (false) {
      // @ts-expect-error presLayout is getter-only
      custom.presLayout = edited;
      // @ts-expect-error presentation layout fields are read-only
      edited.width = inches(1);
    }
  });

  it('creates placeholder identity and materializes empty layout placeholders', async () => {
    const document = PptxDocument.create();
    const layout = document.layouts[0]!;
    const master = document.masters[0]!;
    const types = ['title', 'body', 'pic', 'chart', 'tbl', 'media'] as const;
    const indexes = [100, 0, 4_294_967_294, 103, 104, 105] as const;
    const prompts = types.map((type, ordinal) => layout.addPlaceholder(
      `Prompt ${type}`,
      {
        name: `${type}_box`,
        type,
        ...(ordinal === 0 ? {} : { index: indexes[ordinal] }),
        x: inches(1 + ordinal),
        y: inches(2 + ordinal),
        width: inches(3 + ordinal),
        height: inches(4 + ordinal),
      },
    ));
    expect(prompts.map(({ placeholder }) => placeholder)).toEqual(
      types.map((type, ordinal) => ({ type, index: indexes[ordinal] })),
    );
    expect(prompts.every(({ placeholder }) => Object.isFrozen(placeholder))).toBe(true);
    expect(layout.placeholders).toEqual(prompts);
    expect(prompts[0]!.text).toBe('Prompt title');

    const masterPrompt = master.addPlaceholder([{
      runs: [{ text: 'Rich master prompt', style: { bold: true } }],
    }], {
      name: 'master_title',
      type: 'title',
      index: 200,
    });
    expect(master.placeholders).toEqual([masterPrompt]);
    expect(masterPrompt.richText[0]?.runs[0]).toMatchObject({
      text: 'Rich master prompt',
      style: { bold: true },
    });
    const ordinary = layout.addText('Inherited ordinary object', { name: 'ordinary_layout' });
    expect(layout.shapes).toContain(ordinary);

    const beforeInvalid = document.opcPackage.mutations.map((mutation) => ({ ...mutation }));
    expect(() => layout.addPlaceholder('Duplicate name', {
      name: 'title_box',
      type: 'body',
      index: 300,
    })).toThrow(/name/i);
    expect(() => layout.addPlaceholder('Duplicate identity', {
      name: 'other_title',
      type: 'title',
      index: 100,
    })).toThrow(/identity/i);
    expect(() => layout.addPlaceholder('Reserved index', {
      name: 'reserved',
      type: 'body',
      index: 4_294_967_295,
    })).toThrow(/index/i);
    expect(document.opcPackage.mutations).toEqual(beforeInvalid);

    const slide = document.addSlide({ masterName: 'DEFAULT' });
    expect(slide.shapes.find(({ name }) => name === 'ordinary_layout')).toBeUndefined();
    const materialized = prompts.map(({ name }) =>
      slide.shapes.find((shape) => shape.name === name)!);
    expect(materialized.map(({ placeholder }) => placeholder)).toEqual(
      prompts.map(({ placeholder }) => placeholder),
    );
    expect(materialized.map(({ transform }) => transform)).toEqual(
      prompts.map(({ transform }) => transform),
    );
    for (const shape of materialized) {
      expect(shape).toBeInstanceOf(ShapeModel);
      expect((shape as ShapeModel).richText.flatMap(({ runs }) => runs)).toEqual([]);
    }

    const reopened = await PptxDocument.open(await document.write());
    const reopenedLayout = reopened.layouts[0]!;
    expect(reopenedLayout.placeholders.map(({ placeholder }) => placeholder))
      .toEqual(prompts.map(({ placeholder }) => placeholder));
    expect(reopened.slides.at(-1)?.shapes.map(({ placeholder }) => placeholder))
      .toEqual(prompts.map(({ placeholder }) => placeholder));
  });

  it('materializes alternate-prefix placeholders and ignores foreign or unknown identities', () => {
    const document = PptxDocument.create();
    const layout = document.layouts[0]!;
    document.addSlide({ masterName: 'DEFAULT' });
    const part = document.opcPackage.requirePart(layout.partUri);
    const source = new TextDecoder().decode(part.bytes);
    const alternate = '<q:sp xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main" '
      + 'xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main">'
      + '<q:nvSpPr><q:cNvPr id="20" name="alternate_title"/><q:cNvSpPr/>'
      + '<q:nvPr><q:ph type="title" idx="20"/></q:nvPr></q:nvSpPr><q:spPr>'
      + '<d:xfrm rot="60000"><d:off x="100" y="200"/><d:ext cx="300" cy="400"/>'
      + '</d:xfrm></q:spPr><q:txBody><d:bodyPr anchor="b"/><d:lstStyle/>'
      + '<d:p><d:r><d:t>Prompt</d:t></d:r></d:p></q:txBody></q:sp>';
    const foreign = '<p:sp xmlns:x="urn:foreign"><p:nvSpPr><p:cNvPr id="21" '
      + 'name="foreign_title"/><p:cNvSpPr/><p:nvPr><x:ph type="title" idx="21"/>'
      + '</p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>'
      + '<a:p/></p:txBody></p:sp>';
    const unknown = '<p:sp><p:nvSpPr><p:cNvPr id="22" name="unknown_title"/>'
      + '<p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle" idx="22"/></p:nvPr>'
      + '</p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/>'
      + '</p:txBody></p:sp>';
    document.opcPackage.setPart(
      layout.partUri,
      source.replace('</p:spTree>', `${alternate}${foreign}${unknown}</p:spTree>`),
      part.contentType,
    );

    const beforeRead = document.opcPackage.mutations.map((mutation) => ({ ...mutation }));
    const prompt = layout.shapes.find(({ name }) => name === 'alternate_title')!;
    expect(prompt.placeholder).toEqual({ type: 'title', index: 20 });
    expect(layout.placeholders).toEqual([prompt]);
    expect(layout.placeholders).toEqual([prompt]);
    expect(document.opcPackage.mutations).toEqual(beforeRead);

    const slide = document.addSlide();
    const materialized = slide.shapes.find(({ name }) => name === 'alternate_title')!;
    expect(materialized.placeholder).toEqual({ type: 'title', index: 20 });
    expect(materialized.transform).toMatchObject({
      x: 100,
      y: 200,
      width: 300,
      height: 400,
      rotation: 60_000,
    });
    expect((materialized as ShapeModel).richText.flatMap(({ runs }) => runs)).toEqual([]);
    expect(slide.shapes.find(({ name }) => name === 'foreign_title')).toBeUndefined();
    expect(slide.shapes.find(({ name }) => name === 'unknown_title')).toBeUndefined();
    expect(new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes))
      .toContain('<d:bodyPr anchor="b"/>');
  });

  it('materializes a non-shape layout placeholder without forging text box state', () => {
    const document = PptxDocument.create();
    const layout = document.layouts[0]!;
    const part = document.opcPackage.requirePart(layout.partUri);
    const source = new TextDecoder().decode(part.bytes);
    const picturePlaceholder = '<p:pic><p:nvPicPr><p:cNvPr id="20" '
      + 'name="picture_owner"/><p:cNvPicPr/><p:nvPr><p:ph type="pic" idx="20"/>'
      + '</p:nvPr></p:nvPicPr><p:blipFill/><p:spPr><a:xfrm><a:off x="1" y="2"/>'
      + '<a:ext cx="3" cy="4"/></a:xfrm></p:spPr></p:pic>';
    document.opcPackage.setPart(
      layout.partUri,
      source.replace('</p:spTree>', `${picturePlaceholder}</p:spTree>`),
      part.contentType,
    );

    const slide = document.addSlide();
    const materialized = slide.placeholders.find(({ name }) => name === 'picture_owner')!;
    expect(materialized).toBeInstanceOf(ShapeModel);
    expect((materialized as ShapeModel).isTextBox).toBe(false);
    expect(materialized.placeholder).toEqual({ type: 'pic', index: 20 });
    expect(materialized.transform).toMatchObject({ x: 1, y: 2, width: 3, height: 4 });
    const slideXml = new TextDecoder().decode(
      document.opcPackage.requirePart(slide.partUri).bytes,
    );
    expect(slideXml).toContain(
      '<p:cNvPr id="2" name="picture_owner"/><p:cNvSpPr/><p:nvPr>',
    );
  });

  it('round-trips empty layout placeholders in all six presentation formats', async () => {
    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const document = PptxDocument.create({ format });
      const prompt = document.layouts[0]!.addPlaceholder('Format prompt', {
        name: 'format_body',
        type: 'body',
        index: 7,
        x: inches(1),
        y: inches(2),
        width: inches(3),
        height: inches(4),
      });
      const slide = document.addSlide({ masterName: 'DEFAULT' });
      expect(slide.placeholders[0]?.placeholder, format).toEqual(prompt.placeholder);

      const reopened = await PptxDocument.open(await document.write());
      expect(reopened.format, format).toBe(format);
      expect(reopened.layouts[0]?.placeholders[0]?.placeholder, format)
        .toEqual({ type: 'body', index: 7 });
      expect(reopened.slides.at(-1)?.placeholders[0]?.placeholder, format)
        .toEqual({ type: 'body', index: 7 });
      expect((reopened.slides.at(-1)?.placeholders[0] as ShapeModel).richText
        .flatMap(({ runs }) => runs), format).toEqual([]);
    }
  });

  it('rolls back addSlide when a layout placeholder set is unsafe', async () => {
    const cases = [
      '<p:sp><p:nvSpPr><p:cNvPr id="20" name="ambiguous"/>'
        + '<p:cNvSpPr/><p:nvPr><p:ph type="title" idx="1"/>'
        + '<p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr><p:spPr/>'
        + '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>',
      '<p:sp><p:nvSpPr><p:cNvPr id="20" name="unsupported"/>'
        + '<p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle" idx="1"/>'
        + '</p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>'
        + '<a:p/></p:txBody></p:sp>',
      '<p:sp><p:nvSpPr><p:cNvPr id="20" name="invalid_txbox"/>'
        + '<p:cNvSpPr txBox="maybe"/><p:nvPr><p:ph type="title" idx="1"/>'
        + '</p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>'
        + '<a:p/></p:txBody></p:sp>',
      '<p:sp><p:nvSpPr><p:cNvPr id="20" name="ambiguous_txbox"/>'
        + '<p:cNvSpPr txBox="1" txBox="0"/><p:nvPr><p:ph type="title" idx="1"/>'
        + '</p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>'
        + '<a:p/></p:txBody></p:sp>',
    ];
    for (const unsafe of cases) {
      const document = PptxDocument.create();
      const layout = document.layouts[0]!;
      const part = document.opcPackage.requirePart(layout.partUri);
      const source = new TextDecoder().decode(part.bytes);
      document.opcPackage.setPart(
        layout.partUri,
        source.replace('</p:spTree>', `${unsafe}</p:spTree>`),
        part.contentType,
      );
      const { output: _beforeOutput, ...before } = await sdkPackageSnapshot(document) as {
        readonly output: Uint8Array;
        readonly [key: string]: unknown;
      };

      expect(() => document.addSlide({ masterName: 'DEFAULT' })).toThrow(/placeholder/i);
      const { output: _afterOutput, ...after } = await sdkPackageSnapshot(document) as {
        readonly output: Uint8Array;
        readonly [key: string]: unknown;
      };
      expect(after).toEqual(before);
    }
  });

  it('populate text shape placeholder owners in place with layout geometry', async () => {
    const document = PptxDocument.create();
    const layout = document.layouts[0]!;
    const titlePrompt = layout.addPlaceholder('Title prompt', {
      name: 'title_box',
      type: 'title',
      index: 103,
      x: inches(1),
      y: inches(2),
      width: inches(3),
      height: inches(1),
    });
    const richPrompt = layout.addPlaceholder('Rich prompt', {
      name: 'rich_box',
      type: 'body',
      index: 104,
      x: inches(1),
      y: inches(3),
      width: inches(4),
      height: inches(2),
    });
    const rectPrompt = layout.addPlaceholder('Rect prompt', {
      name: 'rect_box',
      type: 'body',
      index: 105,
      x: inches(5),
      y: inches(1),
      width: inches(2),
      height: inches(2),
    });
    const linePrompt = layout.addPlaceholder('Line prompt', {
      name: 'line_box',
      type: 'body',
      index: 106,
      x: inches(5),
      y: inches(4),
      width: inches(2),
      height: inches(1),
    });
    layout.addPlaceholder('Picture prompt', {
      name: 'pic_box',
      type: 'pic',
      index: 107,
    });
    const slide = document.addSlide({ masterName: 'DEFAULT' });
    const titleEmpty = slide.placeholders.find(({ name }) => name === 'title_box')!;
    const richEmpty = slide.placeholders.find(({ name }) => name === 'rich_box')!;
    const rectEmpty = slide.placeholders.find(({ name }) => name === 'rect_box')!;
    const lineEmpty = slide.placeholders.find(({ name }) => name === 'line_box')!;

    const title = slide.addText('Filled title', {
      placeholder: 'title_box',
      x: inches(9),
      y: inches(9),
      width: inches(9),
      height: inches(9),
    });
    expect(title.id).toBe(titleEmpty.id);
    expect(title.name).toBe('title_box');
    expect(title.placeholder).toEqual({ type: 'title', index: 103 });
    expect(title.transform).toEqual(titlePrompt.transform);
    expect(title.richText[0]?.runs[0]?.text).toBe('Filled title');
    expect(() => titleEmpty.transform).toThrow(/stale/i);

    const rich = slide.addRichText([{
      runs: [{ text: 'Filled rich', style: { bold: true } }],
    }], {
      placeholder: { type: 'body', index: 104 },
    });
    expect(rich.id).toBe(richEmpty.id);
    expect(rich.placeholder).toEqual({ type: 'body', index: 104 });
    expect(rich.transform).toEqual(richPrompt.transform);
    expect(rich.richText[0]?.runs[0]).toMatchObject({
      text: 'Filled rich',
      style: { bold: true },
    });

    const rect = slide.addShape('rect', {
      placeholder: 'rect_box',
      x: inches(9),
      y: inches(9),
      width: inches(9),
      height: inches(9),
      fill: { kind: 'solid', color: { kind: 'srgb', value: '4472C4' } },
    });
    expect(rect.id).toBe(rectEmpty.id);
    expect(rect.placeholder).toEqual({ type: 'body', index: 105 });
    expect(rect.transform).toEqual(rectPrompt.transform);

    const line = slide.addShape('line', {
      placeholder: { type: 'body', index: 106 },
      line: { kind: 'line', color: { kind: 'srgb', value: 'FF3399' } },
    });
    expect(line.id).toBe(lineEmpty.id);
    expect(line.placeholder).toEqual({ type: 'body', index: 106 });
    expect(line.transform).toEqual(linePrompt.transform);
    expect(slide.placeholders.find(({ name }) => name === 'pic_box')).toBeDefined();
    expect(slide.placeholders.map(({ name }) => name)).toEqual([
      'title_box',
      'rich_box',
      'rect_box',
      'line_box',
      'pic_box',
    ]);

    const { output: _beforeOutput, ...before } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(() => slide.addText('Unknown', { placeholder: 'missing' })).toThrow(/placeholder/i);
    expect(() => slide.addText('Second fill', { placeholder: 'title_box' }))
      .toThrow(/empty|filled/i);
    expect(() => slide.addText('Wrong domain', { placeholder: 'pic_box' }))
      .toThrow(/domain|type/i);
    const { output: _afterOutput, ...after } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(after).toEqual(before);

    const duplicate = document.duplicateSlide(document.slides.indexOf(slide));
    document.moveSlide(document.slides.indexOf(duplicate), 0);
    expect(duplicate.placeholders.find(({ name }) => name === 'title_box')).toMatchObject({
      id: title.id,
      placeholder: { type: 'title', index: 103 },
    });
    document.deleteSlide(document.slides.indexOf(slide));
    expect(document.slides).toContain(duplicate);
    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides[0]?.placeholders.find(({ name }) => name === 'title_box'))
      .toMatchObject({
        id: title.id,
        placeholder: { type: 'title', index: 103 },
      });

    const malformed = document.addSlide({ masterName: 'DEFAULT' });
    const layoutRelationship = malformed.relationships.find(({ type }) =>
      type.endsWith('/slideLayout'))!;
    document.opcPackage.addRelationship(malformed.partUri, {
      type: layoutRelationship.type,
      target: layoutRelationship.target,
      targetMode: 'Internal',
    });
    const malformedBefore = document.opcPackage.mutations.map((mutation) => ({ ...mutation }));
    expect(() => malformed.addShape('line', { placeholder: 'line_box' }))
      .toThrow(/layout relationship/i);
    expect(document.opcPackage.mutations).toEqual(malformedBefore);

    const missingOwner = document.addSlide({ masterName: 'DEFAULT' });
    const missingShape = missingOwner.placeholders.find(({ name }) => name === 'line_box')!;
    const { xml: missingXml, element: missingElement } = missingOwner.resolveShape(missingShape.id);
    missingXml.removeElement(missingElement);
    document.opcPackage.setPart(
      missingOwner.partUri,
      missingXml.serialize(),
      document.opcPackage.requirePart(missingOwner.partUri).contentType,
    );
    const missingBefore = document.opcPackage.mutations.map((mutation) => ({ ...mutation }));
    expect(() => missingOwner.addShape('line', { placeholder: 'line_box' }))
      .toThrow(/owner.*missing|missing.*owner/i);
    expect(document.opcPackage.mutations).toEqual(missingBefore);

    const ambiguousDocument = PptxDocument.create();
    const ambiguousLayout = ambiguousDocument.layouts[0]!;
    const ambiguousPrompt = ambiguousLayout.addPlaceholder('Ambiguous', {
      name: 'ambiguous_box',
      type: 'body',
      index: 10,
    });
    const ambiguousSlide = ambiguousDocument.addSlide({ masterName: 'DEFAULT' });
    const ambiguousPart = ambiguousDocument.opcPackage.requirePart(ambiguousLayout.partUri);
    const ambiguousXml = LosslessXmlDocument.parse(ambiguousPart.bytes);
    const ambiguousElement = ambiguousXml.elements('sp').find((shape) => {
      const properties = ambiguousXml.descendants(shape, 'cNvPr')[0];
      return ambiguousXml.attribute(properties!, 'id')?.value === String(ambiguousPrompt.id);
    })!;
    const clone = ambiguousXml.original(ambiguousElement)
      .replace(`id="${ambiguousPrompt.id}"`, `id="${ambiguousPrompt.id + 1}"`)
      .replace('idx="10"', 'idx="11"');
    const tree = ambiguousXml.elements('spTree')[0]!;
    ambiguousXml.appendChildXml(tree, clone);
    ambiguousDocument.opcPackage.setPart(
      ambiguousLayout.partUri,
      ambiguousXml.serialize(),
      ambiguousPart.contentType,
    );
    const ambiguousBefore = ambiguousDocument.opcPackage.mutations
      .map((mutation) => ({ ...mutation }));
    expect(() => ambiguousSlide.addText('Ambiguous', { placeholder: 'ambiguous_box' }))
      .toThrow(/duplicate|ambiguous/i);
    expect(ambiguousDocument.opcPackage.mutations).toEqual(ambiguousBefore);
  });

  it('populate image chart placeholders from high-level sources and chart groups', async () => {
    const document = PptxDocument.create();
    const layout = document.layouts[0]!;
    const rasterPrompt = layout.addPlaceholder('Raster prompt', {
      name: 'raster_image',
      type: 'pic',
      index: 201,
      x: inches(1),
      y: inches(1),
      width: inches(3),
      height: inches(2),
    });
    const svgPrompt = layout.addPlaceholder('SVG prompt', {
      name: 'vector_image',
      type: 'pic',
      index: 202,
      x: inches(4),
      y: inches(1),
      width: inches(2),
      height: inches(2),
    });
    const singlePrompt = layout.addPlaceholder('Single chart prompt', {
      name: 'single_chart',
      type: 'chart',
      index: 203,
      x: inches(1),
      y: inches(3),
      width: inches(3),
      height: inches(2),
    });
    const comboPrompt = layout.addPlaceholder('Combo chart prompt', {
      name: 'combo_chart',
      type: 'chart',
      index: 204,
      x: inches(4),
      y: inches(3),
      width: inches(3),
      height: inches(2),
    });
    const slide = document.addSlide({ masterName: 'DEFAULT' });
    const empty = [...slide.shapes];

    const raster = await document.addImage(0, sdkPngDataUri(sdkPngHeader(16, 9)), {
      placeholder: 'raster_image',
      sizing: { type: 'cover', width: inches(9), height: inches(9) },
    });
    const vector = await document.addImage(
      0,
      new Blob([sdkSvg(640, 360)], { type: 'image/svg+xml' }),
      {
        placeholder: { type: 'pic', index: 202 },
        fallback: new Blob([sdkPngHeader(1, 1)], { type: 'image/png' }),
        width: inches(9),
        height: inches(9),
      },
    );
    const single = await document.addChart(0, 'bar', [{
      name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20],
    }], {
      placeholder: 'single_chart',
      x: inches(9),
      width: inches(9),
    });
    const combo = await document.addChart(0, [
      {
        type: 'bar',
        series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
      },
      {
        type: 'line',
        axis: 'secondary',
        series: [{ name: 'Trend', categories: ['Q1', 'Q2'], values: [11, 21] }],
      },
    ], {
      placeholder: { type: 'chart', index: 204 },
      height: inches(9),
    });

    expect([raster.id, vector.id, single.id, combo.id]).toEqual(empty.map(({ id }) => id));
    expect([raster.name, vector.name, single.name, combo.name]).toEqual([
      'raster_image',
      'vector_image',
      'single_chart',
      'combo_chart',
    ]);
    expect([raster.placeholder, vector.placeholder, single.placeholder, combo.placeholder]).toEqual([
      { type: 'pic', index: 201 },
      { type: 'pic', index: 202 },
      { type: 'chart', index: 203 },
      { type: 'chart', index: 204 },
    ]);
    expect([raster.transform, vector.transform, single.transform, combo.transform]).toEqual([
      rasterPrompt.transform,
      svgPrompt.transform,
      singlePrompt.transform,
      comboPrompt.transform,
    ]);
    expect(raster.sourceRectangle).toEqual({
      left: 7.813,
      top: 0,
      right: 7.813,
      bottom: 0,
    });
    expect(vector.isSvg).toBe(true);
    expect(vector.fallbackPartUri).toBeDefined();
    expect(vector.svgPartUri).toBeDefined();
    expect(single.definition?.groups.map(({ type }) => type)).toEqual(['bar']);
    expect(combo.definition?.groups.map(({ type, axis }) => [type, axis])).toEqual([
      ['bar', 'primary'],
      ['line', 'secondary'],
    ]);
    for (const owner of empty) expect(() => owner.name).toThrow(/stale/i);
    expect(slide.shapes).toEqual([raster, vector, single, combo]);
    expect(slide.placeholders).toEqual([raster, vector, single, combo]);
    const source = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(source.match(/<p:ph type="pic" idx="20[12]"\/>/g)).toHaveLength(2);
    expect(source.match(/<p:ph type="chart" idx="20[34]"\/>/g)).toHaveLength(2);

    const { output: _beforeOutput, ...before } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    await expect(document.addImage(0, sdkPngHeader(1, 1), {
      placeholder: 'raster_image',
    })).rejects.toThrow(/empty|filled/i);
    await expect(document.addChart(0, 'bar', [{
      name: 'Wrong domain', categories: ['Q1'], values: [1],
    }], { placeholder: 'vector_image' })).rejects.toThrow(/domain|type/i);
    const { output: _afterOutput, ...after } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(after).toEqual(before);

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides[0]?.shapes.map(({ placeholder }) => placeholder)).toEqual([
      { type: 'pic', index: 201 },
      { type: 'pic', index: 202 },
      { type: 'chart', index: 203 },
      { type: 'chart', index: 204 },
    ]);
    expect((reopened.slides[0]?.shapes[1] as ImageModel).isSvg).toBe(true);
    expect((reopened.slides[0]?.shapes[3] as ChartModel).definition?.groups).toHaveLength(2);

    const duplicate = document.duplicateSlide(0);
    expect(duplicate.shapes.map(({ placeholder }) => placeholder)).toEqual(
      slide.shapes.map(({ placeholder }) => placeholder),
    );
    document.deleteSlide(document.slides.indexOf(slide));
    expect(document.opcPackage.parts.some(({ uri }) => uri.startsWith('/ppt/media/'))).toBe(true);
    document.deleteSlide(document.slides.indexOf(duplicate));
    expect(document.opcPackage.parts.filter(({ uri }) =>
      uri.startsWith('/ppt/charts/') || uri.startsWith('/ppt/embeddings/'))).toEqual([]);
    expect(document.opcPackage.parts.filter(({ uri }) => uri.startsWith('/ppt/media/')))
      .toHaveLength(3);
    expect(document.opcPackage.graph.filter(({ uri }) => uri.startsWith('/ppt/media/'))
      .every(({ incoming }) => incoming.length === 0)).toBe(true);
  });

  it('populate table media placeholders with native state and inherited geometry', async () => {
    const document = PptxDocument.create();
    const layout = document.layouts[0]!;
    const tablePrompt = layout.addPlaceholder('Table prompt', {
      name: 'data_table',
      type: 'tbl',
      index: 205,
      x: inches(1),
      y: inches(1),
      width: inches(4),
      height: inches(2),
      rotation: degrees(15),
      flipHorizontal: true,
    });
    const audioPrompt = layout.addPlaceholder('Audio prompt', {
      name: 'narration',
      type: 'media',
      index: 206,
      x: inches(1),
      y: inches(3.5),
      width: inches(2),
      height: inches(1),
      rotation: degrees(-10),
      flipVertical: true,
    });
    const videoPrompt = layout.addPlaceholder('Video prompt', {
      name: 'demo_video',
      type: 'media',
      index: 207,
      x: inches(4),
      y: inches(3.5),
      width: inches(3),
      height: inches(2),
    });
    const externalPrompt = layout.addPlaceholder('External prompt', {
      name: 'external_video',
      type: 'media',
      index: 208,
      x: inches(7.5),
      y: inches(3.5),
      width: inches(2),
      height: inches(2),
    });
    const slide = document.addSlide({ masterName: 'DEFAULT' });
    const empty = [...slide.shapes];

    const table = slide.addTable([
      [
        {
          text: 'Quarter',
          options: {
            fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
            valign: 'middle',
          },
        },
        { text: 'Revenue', options: { align: 'right', fit: 'shrink' } },
      ],
      ['Q1', '10'],
    ], {
      placeholder: 'data_table',
      width: inches(9),
      height: inches(9),
      columnWidths: [inches(3), inches(6)],
      rowHeights: [inches(3), inches(6)],
    });
    const audio = await document.addAudio(0, Uint8Array.of(1, 2, 3), {
      placeholder: { type: 'media', index: 206 },
      contentType: 'audio/mpeg',
      poster: Uint8Array.of(4, 5),
      posterContentType: 'image/png',
      play: 'auto',
      loop: true,
      volume: 0.5,
      x: inches(9),
      width: inches(9),
    });
    const video = await document.addVideo(0, Uint8Array.of(6, 7, 8), {
      placeholder: 'demo_video',
      contentType: 'video/mp4',
      poster: Uint8Array.of(9, 10),
      posterContentType: 'image/gif',
      hideWhenStopped: true,
      y: inches(9),
      height: inches(9),
    });
    const external = await document.addVideo(0, 'https://example.com/demo.mp4', {
      placeholder: 'external_video',
      poster: Uint8Array.of(11, 12),
      posterContentType: 'image/jpeg',
    });

    expect([table.id, audio.id, video.id, external.id]).toEqual(empty.map(({ id }) => id));
    expect([table.name, audio.name, video.name, external.name]).toEqual([
      'data_table',
      'narration',
      'demo_video',
      'external_video',
    ]);
    expect([table.placeholder, audio.placeholder, video.placeholder, external.placeholder])
      .toEqual([
        { type: 'tbl', index: 205 },
        { type: 'media', index: 206 },
        { type: 'media', index: 207 },
        { type: 'media', index: 208 },
      ]);
    expect([table.transform, audio.transform, video.transform, external.transform]).toEqual([
      tablePrompt.transform,
      audioPrompt.transform,
      videoPrompt.transform,
      externalPrompt.transform,
    ]);
    expect(table.rows.map(({ cells }) => cells.map(({ text }) => text))).toEqual([
      ['Quarter', 'Revenue'],
      ['Q1', '10'],
    ]);
    expect(table.rows[0]?.cells[0]).toMatchObject({
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
      verticalAlignment: 'middle',
    });
    expect(table.columnWidths?.reduce((sum, value) => sum + value, 0))
      .toBe(tablePrompt.transform.width);
    expect(table.rowHeights?.reduce((sum, value) => sum + value, 0))
      .toBe(tablePrompt.transform.height);
    expect(audio.settings).toEqual({
      play: 'auto',
      loop: true,
      hideWhenStopped: false,
      volume: 0.5,
    });
    expect(video.settings.hideWhenStopped).toBe(true);
    expect(external.externalUrl).toBe('https://example.com/demo.mp4');
    expect(audio.mediaPartUri).toBeDefined();
    expect(audio.posterPartUri).toBeDefined();
    expect(video.mediaPartUri).toBeDefined();
    expect(video.posterPartUri).toBeDefined();
    for (const owner of empty) expect(() => owner.name).toThrow(/stale/i);
    expect(slide.shapes).toEqual([table, audio, video, external]);
    const source = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(source).toContain('<p:ph type="tbl" idx="205"/>');
    expect(source.match(/<p:ph type="media" idx="20[678]"\/>/g)).toHaveLength(3);
    expect(source).toContain('<p:timing>');

    await video.replacePoster(Uint8Array.of(13, 14), { contentType: 'image/png' });
    const originalAudioUri = audio.mediaPartUri!;
    const originalVideoPosterUri = video.posterPartUri!;
    const duplicate = document.duplicateSlide(0);
    const duplicateAudio = duplicate.media.find(({ name }) => name === 'narration')!;
    const duplicateVideo = duplicate.media.find(({ name }) => name === 'demo_video')!;
    expect(duplicateAudio.mediaPartUri).toBe(originalAudioUri);
    expect(duplicateVideo.posterPartUri).toBe(originalVideoPosterUri);
    await duplicateAudio.replaceSource(Uint8Array.of(15, 16), { contentType: 'audio/mpeg' });
    await duplicateVideo.replacePoster(Uint8Array.of(17, 18), { contentType: 'image/gif' });
    expect(duplicateAudio.mediaPartUri).not.toBe(originalAudioUri);
    expect(duplicateVideo.posterPartUri).not.toBe(originalVideoPosterUri);
    expect(audio.mediaPartUri).toBe(originalAudioUri);
    expect(video.posterPartUri).toBe(originalVideoPosterUri);

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides[0]?.shapes.map(({ placeholder }) => placeholder)).toEqual([
      { type: 'tbl', index: 205 },
      { type: 'media', index: 206 },
      { type: 'media', index: 207 },
      { type: 'media', index: 208 },
    ]);
    expect((reopened.slides[0]?.shapes[0] as TableModel).rows[0]?.cells[0]?.text)
      .toBe('Quarter');
    expect(reopened.slides[0]?.media[0]?.settings.play).toBe('auto');

    const { output: _beforeOutput, ...before } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(() => slide.addTable([['Second']], { placeholder: 'data_table' }))
      .toThrow(/empty|filled/i);
    await expect(document.addAudio(0, Uint8Array.of(1), {
      placeholder: 'data_table',
      contentType: 'audio/mpeg',
    })).rejects.toThrow(/domain|type/i);
    const { output: _afterOutput, ...after } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(after).toEqual(before);

    duplicateAudio.remove();
    expect(document.opcPackage.hasPart(originalAudioUri)).toBe(true);
    audio.remove();
    expect(document.opcPackage.hasPart(originalAudioUri)).toBe(false);

    const rollback = PptxDocument.create();
    rollback.layouts[0]!.addPlaceholder('Rollback media', {
      name: 'rollback_media',
      type: 'media',
      index: 209,
    });
    const rollbackSlide = rollback.addSlide({ masterName: 'DEFAULT' });
    const { output: _rollbackOutput, ...rollbackBefore } = await sdkPackageSnapshot(rollback) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    const originalSetPart = rollback.opcPackage.setPart.bind(rollback.opcPackage);
    const setPart = vi.spyOn(rollback.opcPackage, 'setPart')
      .mockImplementation((uri, bytes, contentType) => {
        if (uri === rollbackSlide.partUri) throw new Error('placeholder media write failed');
        return originalSetPart(uri, bytes, contentType);
      });
    await expect(rollback.addVideo(0, Uint8Array.of(1, 2), {
      placeholder: 'rollback_media',
      contentType: 'video/mp4',
    })).rejects.toThrow('placeholder media write failed');
    setPart.mockRestore();
    const { output: _rollbackAfterOutput, ...rollbackAfter } = await sdkPackageSnapshot(rollback) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(rollbackAfter).toEqual(rollbackBefore);
  });

  it('define slide master with synchronous objects, margin, and slide numbers', async () => {
    const document = PptxDocument.create({ firstSlideNumber: 7 });
    const master = document.masters[0]!;
    const layout = await document.defineSlideMaster({
      title: 'BRAND',
      master,
      background: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 10,
      },
      margin: [inches(0.1), inches(0.2), inches(0.3), inches(0.4)],
      slideNumber: { x: 800, y: 500, width: 100, height: 30 },
      objects: [
        {
          kind: 'rect',
          options: { x: inches(0.1), y: inches(0.2), width: inches(0.3), height: inches(0.4) },
        },
        {
          kind: 'line',
          options: { x: inches(0.5), y: inches(0.6), width: inches(0.7), height: inches(0.01) },
        },
        { kind: 'text', text: 'Brand', options: { x: inches(0.8), y: inches(0.9) } },
        {
          kind: 'placeholder',
          text: 'Title prompt',
          options: { name: 'title_box', type: 'title', x: inches(1), y: inches(1.1) },
        },
      ],
    });

    expect(layout).toBe(document.layouts.find(({ partUri }) => partUri === layout.partUri));
    expect(master.layouts.at(-1)).toBe(layout);
    expect(layout.name).toBe('BRAND');
    expect(layout.masterPartUri).toBe(master.partUri);
    expect(layout.margin).toEqual({
      top: inches(0.1),
      right: inches(0.2),
      bottom: inches(0.3),
      left: inches(0.4),
    });
    expect(Object.isFrozen(layout.margin)).toBe(true);
    expect(layout.background).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 10,
    });
    expect(layout.shapes.map(({ kind }) => kind)).toEqual([
      'shape',
      'shape',
      'text',
      'text',
      'text',
    ]);
    expect(layout.placeholders.map(({ name, placeholder }) => ({ name, placeholder })))
      .toEqual([{ name: 'title_box', placeholder: { type: 'title', index: 103 } }]);
    expect(layout.slideNumber).toMatchObject({ x: 800, y: 500, width: 100, height: 30 });
    expect(new Set(layout.shapes.map(({ id }) => id)).size).toBe(layout.shapes.length);
    const masterSource = new TextDecoder().decode(
      document.opcPackage.requirePart(master.partUri).bytes,
    );
    expect(masterSource).toMatch(/<p:hf[^>]*sldNum="1"/);

    const slide = document.addSlide({ masterName: 'BRAND' });
    expect(slide.shapes.map(({ kind }) => kind)).toEqual(['text', 'text']);
    expect(slide.placeholders[0]?.placeholder).toEqual({ type: 'title', index: 103 });
    expect(slide.slideNumber).toMatchObject({ x: 800, y: 500, width: 100, height: 30 });
    const slideSource = new TextDecoder().decode(
      document.opcPackage.requirePart(slide.partUri).bytes,
    );
    expect(slideSource).toContain('<a:t>7</a:t>');

    const reopened = await PptxDocument.open(await document.write());
    const reopenedLayout = reopened.layouts.find(({ name }) => name === 'BRAND')!;
    expect(reopenedLayout.margin).toBeUndefined();
    expect(reopenedLayout.shapes.map(({ kind }) => kind)).toEqual(layout.shapes.map(({ kind }) => kind));
    expect(reopened.slides[0]?.slideNumber).toMatchObject({ x: 800, y: 500 });
  });

  it('define slide master rejects unsafe definitions without observable mutation', async () => {
    const document = PptxDocument.create();
    const foreignMaster = PptxDocument.create().masters[0]!;
    const { output: _beforeOutput, ...before } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    let titleReads = 0;
    const accessor = Object.defineProperty({}, 'title', {
      enumerable: true,
      get: () => {
        titleReads += 1;
        return 'ACCESSOR';
      },
    });
    const sparse: unknown[] = [];
    sparse.length = 1;
    const invalid = [
      null,
      accessor,
      { title: '' },
      { title: 'BAD\u0000NAME' },
      { title: 'UNKNOWN', unknown: true },
      { title: 'SYMBOL', [Symbol('unsafe')]: true },
      { title: 'FOREIGN', master: foreignMaster },
      {
        title: 'MARGIN-X',
        margin: [inches(0.1), document.slideSize.width, inches(0.1), inches(0.01)],
      },
      {
        title: 'MARGIN-Y',
        margin: [document.slideSize.height, inches(0.01), inches(0.01), inches(0.01)],
      },
      { title: 'SPARSE', objects: sparse },
      { title: 'OBJECT', objects: [{ kind: 'unknown' }] },
      { title: 'SLIDE-NUMBER', slideNumber: { width: 0 } },
      {
        title: 'BACKGROUND',
        background: { kind: 'solid', color: { kind: 'srgb', value: 'FFF' } },
      },
      { title: 'BACKGROUND-SOURCE', background: { kind: 'image-source' } },
      {
        title: 'BACKGROUND-CONTENT-TYPE',
        background: {
          kind: 'image-source',
          source: sdkPngHeader(1, 1),
          contentType: 'image/svg+xml',
        },
      },
      { title: 'RECT', objects: [{ kind: 'rect', options: { width: 0 } }] },
      { title: 'TEXT', objects: [{ kind: 'text', text: 1 }] },
      {
        title: 'TEXT-OPTION',
        objects: [{ kind: 'text', text: 'Text', options: { unknown: true } }],
      },
      {
        title: 'PLACEHOLDER',
        objects: [{ kind: 'placeholder', options: { name: '', type: 'title' } }],
      },
      { title: 'IMAGE-SOURCE', objects: [{ kind: 'image' }] },
      { title: 'CHART-GROUPS', objects: [{ kind: 'chart' }] },
    ];
    for (const definition of invalid) {
      await expect(document.defineSlideMaster(definition as never)).rejects.toThrow();
      const { output: _afterOutput, ...after } = await sdkPackageSnapshot(document) as {
        readonly output: Uint8Array;
        readonly [key: string]: unknown;
      };
      expect(after).toEqual(before);
    }
    expect(titleReads).toBe(0);

    const created = await document.defineSlideMaster({ title: 'UNIQUE' });
    expect(created.margin).toBeUndefined();
    const { output: _createdOutput, ...afterCreated } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    await expect(document.defineSlideMaster({ title: 'UNIQUE' })).rejects.toThrow(/unique|already/i);
    const { output: _duplicateOutput, ...afterDuplicate } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(afterDuplicate).toEqual(afterCreated);

    const rollback = PptxDocument.create();
    const { output: _rollbackOutput, ...rollbackBefore } = await sdkPackageSnapshot(rollback) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    const originalSetPart = rollback.opcPackage.setPart.bind(rollback.opcPackage);
    let layoutWrites = 0;
    const setPart = vi.spyOn(rollback.opcPackage, 'setPart')
      .mockImplementation((uri, bytes, contentType) => {
        if (uri === '/ppt/slideLayouts/slideLayout2.xml') {
          layoutWrites += 1;
          if (layoutWrites === 2) throw new Error('define slide master write failed');
        }
        return originalSetPart(uri, bytes, contentType);
      });
    await expect(rollback.defineSlideMaster({
      title: 'ROLLBACK',
      background: { kind: 'solid', color: { kind: 'srgb', value: '112233' } },
      margin: inches(0.5),
    })).rejects.toThrow('define slide master write failed');
    setPart.mockRestore();
    const { output: _rollbackAfterOutput, ...rollbackAfter } = await sdkPackageSnapshot(rollback) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(rollbackAfter).toEqual(rollbackBefore);
    const retried = await rollback.defineSlideMaster({
      title: 'ROLLBACK',
      margin: inches(0.5),
    });
    expect(retried.margin).toEqual({
      top: inches(0.5),
      right: inches(0.5),
      bottom: inches(0.5),
      left: inches(0.5),
    });
  });

  it('define slide master works in all six presentation formats', async () => {
    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const document = PptxDocument.create({ format, firstSlideNumber: 2 });
      const title = `BRAND-${format}`;
      const layout = await document.defineSlideMaster({
        title,
        margin: inches(0.25),
        slideNumber: { align: 'center' },
        objects: [{
          kind: 'placeholder',
          options: { name: `title_${format}`, type: 'title' },
        }],
      });
      const slide = document.addSlide({ masterName: title });
      const reopened = await PptxDocument.open(await document.write());
      expect(reopened.format).toBe(format);
      expect(layout.margin).toEqual({
        top: inches(0.25),
        right: inches(0.25),
        bottom: inches(0.25),
        left: inches(0.25),
      });
      expect(reopened.layouts.find(({ name }) => name === title)).toBeDefined();
      expect(reopened.slides[0]?.partUri).toBe(slide.partUri);
      expect(reopened.slides[0]?.slideNumber?.align).toBe('center');
    }
  });

  it('reports transient slide master layout margins as profile-specific info only', async () => {
    const document = PptxDocument.create();
    const layout = await document.defineSlideMaster({
      title: 'TRANSIENT-MARGIN',
      margin: inches(0.5),
    });
    for (const compatibility of [
      'powerpoint-2010',
      'powerpoint-current',
      'keynote-current',
      'libreoffice-current',
      'google-slides-import',
    ] as const) {
      await expect(document.write({ compatibility })).resolves.toBeInstanceOf(Uint8Array);
      expect(document.diagnostics.filter(({ code }) => code === 'LAYOUT_MARGIN_TRANSIENT'))
        .toEqual([expect.objectContaining({
          severity: 'info',
          partUri: layout.partUri,
          compatibility,
        })]);
      expect(document.diagnostics.filter(({ severity }) => severity !== 'info')).toEqual([]);
    }
    const reopened = await PptxDocument.open(await document.write());
    await reopened.write();
    expect(reopened.layouts.find(({ name }) => name === 'TRANSIENT-MARGIN')?.margin)
      .toBeUndefined();
    expect(reopened.diagnostics.filter(({ code }) => code === 'LAYOUT_MARGIN_TRANSIENT'))
      .toEqual([]);
  });

  it('uses runtime named-layout margins for auto-page and canonical fallback after reopen', async () => {
    const document = PptxDocument.create({
      slideSize: { width: inches(10), height: inches(5) },
    });
    const title = 'AUTO-PAGE-MARGINS';
    const layoutMargin = [
      inches(1),
      inches(0.25),
      inches(1.25),
      inches(0.75),
    ] as const;
    await document.defineSlideMaster({ title, margin: layoutMargin });
    const rows = [['Header'], ['A'], ['B'], ['C'], ['D']] as const;
    const baseOptions: AddTableOptions = {
      autoPage: true,
      autoPageRepeatHeader: true,
      autoPageHeaderRows: 1,
      y: inches(3),
      columnWidths: [inches(4)],
      rowHeights: [
        inches(0.5),
        inches(0.75),
        inches(0.75),
        inches(0.75),
        inches(0.75),
      ],
    };

    const runtimeSource = document.addSlide({ masterName: title });
    runtimeSource.addTable(rows, baseOptions);
    expect(runtimeSource.newAutoPagedSlides).toHaveLength(2);
    const runtimeTables = [runtimeSource, ...runtimeSource.newAutoPagedSlides].map(
      (slide) => slide.shapes.find((shape): shape is TableModel => shape instanceof TableModel)!,
    );
    expect(runtimeTables.map(({ transform }) => transform.y)).toEqual([
      inches(3),
      layoutMargin[0],
      layoutMargin[0],
    ]);
    expect(runtimeTables.map((table) => table.rows.map((row) => row.cells[0]!.text)))
      .toEqual([['Header'], ['Header', 'A', 'B', 'C'], ['Header', 'D']]);

    const explicitMargin: TableAutoPageMarginInput = [
      inches(0.25),
      inches(0.1),
      inches(0.25),
      inches(0.1),
    ];
    const explicitSource = document.addSlide({ masterName: title });
    explicitSource.addTable(rows, { ...baseOptions, slideMargin: explicitMargin });
    expect(explicitSource.newAutoPagedSlides).toHaveLength(1);
    const explicitTable = explicitSource.newAutoPagedSlides[0]!.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    )!;
    expect(explicitTable.transform.y).toBe(explicitMargin[0]);
    expect(explicitTable.rows.map((row) => row.cells[0]!.text))
      .toEqual(['Header', 'B', 'C', 'D']);

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.layouts.find(({ name }) => name === title)?.margin).toBeUndefined();
    const fallbackSource = reopened.addSlide({ masterName: title });
    fallbackSource.addTable(rows, baseOptions);
    expect(fallbackSource.newAutoPagedSlides).toHaveLength(1);
    const fallbackTables = [fallbackSource, ...fallbackSource.newAutoPagedSlides].map(
      (slide) => slide.shapes.find((shape): shape is TableModel => shape instanceof TableModel)!,
    );
    expect(fallbackTables.map(({ transform }) => transform.y))
      .toEqual([inches(3), inches(0.5)]);
    expect(fallbackTables.map((table) => table.rows.map((row) => row.cells[0]!.text)))
      .toEqual([['Header', 'A'], ['Header', 'B', 'C', 'D']]);
  });

  it('exports measured table auto-page contracts through all six presentation formats', async () => {
    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const document = PptxDocument.create({ format });
      const source = document.addSlide();
      const margin: TableAutoPageMarginInput = inches(0.5);
      const weightedCell: AddTableCellOptions = {
        autoPageCharWeight: 1,
        autoPageLineWeight: -1,
        margin: 0,
      };
      const options: AddTableOptions = {
        autoPage: true,
        autoPageCharWeight: -1,
        autoPageLineWeight: 0,
        autoPageRepeatHeader: true,
        autoPageHeaderRows: 1,
        autoPageSlideStartY: inches(0.75),
        slideMargin: margin,
        y: inches(4.5),
        columnWidths: [inches(4)],
        rowHeights: [0, inches(0.75), 0],
      };
      source.addTable([
        ['Header'],
        [{ text: 'A'.repeat(80), options: weightedCell }],
        ['B'],
      ], options);

      const generated: readonly SlideModel[] = source.newAutoPagedSlides;
      expect(generated).toHaveLength(1);
      expect(Object.isFrozen(generated)).toBe(true);
      const tables = [source, ...generated].map((slide) => slide.shapes.find(
        (shape): shape is TableModel => shape instanceof TableModel,
      )!);
      expect(tables.map((table) => table.rows.map((row) => row.cells[0]!.text)))
        .toEqual([['Header'], ['Header', 'A'.repeat(80), 'B']]);
      expect(tables.map(({ transform }) => transform.y))
        .toEqual([inches(4.5), inches(0.75)]);
      expect(tables.every((table) => table.rowHeights?.every((height) => height > 0)))
        .toBe(true);

      const reopened = await PptxDocument.open(await document.write());
      expect(reopened.format).toBe(format);
      expect(reopened.slides).toHaveLength(2);
      expect(reopened.slides.every((slide) => slide.newAutoPagedSlides.length === 0))
        .toBe(true);
      expect(reopened.slides.map((slide) =>
        (slide.shapes[0] as TableModel).rows.map((row) => row.cells[0]!.text)))
        .toEqual([['Header'], ['Header', 'A'.repeat(80), 'B']]);
    }

    const invalid = PptxDocument.create().addSlide();
    expect(() => invalid.addTable([['A']], {
      autoPage: true,
      autoPageCharWeight: 1.001,
    })).toThrow(/weight/i);
    expect(() => invalid.addTable([['A']], {
      autoPage: true,
      autoPageLineWeight: '0' as never,
    })).toThrow(/weight/i);

    const descriptor = Object.getOwnPropertyDescriptor(
      Object.getPrototypeOf(PptxDocument.create().addSlide()),
      'newAutoPagedSlides',
    );
    expect(descriptor?.get).toBeTypeOf('function');
    expect(descriptor?.set).toBeUndefined();

    if (false) {
      const document = PptxDocument.create();
      const slide = document.addSlide();
      const scalar: TableAutoPageMarginInput = inches(0.5);
      const tuple: TableAutoPageMarginInput = [1, 2, 3, 4];
      slide.addTable([['A']], {
        autoPage: true,
        autoPageRepeatHeader: false,
        autoPageSlideStartY: 0,
        slideMargin: scalar,
        rowHeights: [1],
      });
      slide.addTable([['A']], { autoPage: true, slideMargin: tuple, rowHeights: [1] });
      // @ts-expect-error generated slide state is getter-only.
      slide.newAutoPagedSlides = [];
      // @ts-expect-error generated slide state is a readonly collection.
      slide.newAutoPagedSlides.push(slide);
      const stringBoolean: AddTableOptions = {
        // @ts-expect-error autoPage is boolean-only.
        autoPage: 'true',
      };
      const malformedMargin: AddTableOptions = {
        // @ts-expect-error slideMargin tuple has exactly four values.
        slideMargin: [1, 2, 3],
      };
      const legacyHeader: AddTableOptions = {
        // @ts-expect-error legacy PptxGenJS header alias is not supported.
        addHeaderToEach: true,
      };
      const legacyStart: AddTableOptions = {
        // @ts-expect-error legacy PptxGenJS continuation-Y alias is not supported.
        newSlideStartY: 1,
      };
      const boundedTableWeights: readonly AddTableOptions[] = [
        { autoPage: true, autoPageCharWeight: -1, autoPageLineWeight: 1 },
        { autoPage: true, autoPageCharWeight: 0, autoPageLineWeight: 0 },
        { autoPage: true, autoPageCharWeight: 1, autoPageLineWeight: -1 },
      ];
      const boundedCellWeights: readonly AddTableCellOptions[] = [
        { autoPageCharWeight: -1, autoPageLineWeight: 1 },
        { autoPageCharWeight: 0, autoPageLineWeight: 0 },
        { autoPageCharWeight: 1, autoPageLineWeight: -1 },
      ];
      // @ts-expect-error normalized table internals are not SDK exports.
      type HiddenNormalizedTable = import('./index.js').NormalizedTableDefinition;
      // @ts-expect-error measurement internals are not SDK exports.
      type HiddenMeasuredLine = import('./index.js').MeasuredTableLine;
      const invalidCharWeight: AddTableOptions = {
        autoPage: true,
        // @ts-expect-error measurement weights are numeric.
        autoPageCharWeight: '0',
      };
      void [
        stringBoolean,
        malformedMargin,
        legacyHeader,
        legacyStart,
        boundedTableWeights,
        boundedCellWeights,
        invalidCharWeight,
        undefined as unknown as HiddenNormalizedTable,
        undefined as unknown as HiddenMeasuredLine,
      ];
    }
  });

  it('prepares async slide master definition sources and charts before atomic commit', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'pptx-master-async-'));
    const png = sdkPngHeader(16, 9);
    const jpeg = sdkJpeg(12, 8);
    const gif = sdkGif(10, 6);
    const jpegPath = join(directory, 'source.jpg');
    await writeFile(jpegPath, jpeg);
    const fetchSpy = vi.spyOn(globalThis, 'fetch').mockResolvedValue(
      new Response(gif, { status: 200, headers: { 'content-type': 'image/gif' } }),
    );
    const chartTypes = [
      'area', 'bar', 'bar3D', 'bubble', 'doughnut', 'line', 'pie', 'radar', 'scatter',
    ] as const;
    const chartObjects = chartTypes.map((type, index) => ({
      kind: 'chart' as const,
      groups: [{
        type,
        series: type === 'scatter'
          ? [{ name: type, xValues: [1, 2], values: [10, 20] }]
          : type === 'bubble'
            ? [{ name: type, xValues: [1, 2], values: [10, 20], sizes: [5, 6] }]
            : [{ name: type, categories: ['Q1', 'Q2'], values: [10, 20] }],
      }],
      options: { name: `${type}-${index}`, x: inches(index / 10) },
    }));
    const document = PptxDocument.create();
    try {
      const layout = await document.defineSlideMaster({
        title: 'ASYNC',
        background: {
          kind: 'image-source',
          source: png.buffer.slice(0),
          contentType: 'image/png',
        },
        objects: [
          { kind: 'text', text: 'Async brand' },
          {
            kind: 'image',
            source: sdkPngDataUri(png),
            options: {
              name: 'data-png',
              sizing: { type: 'cover', width: inches(2), height: inches(2) },
            },
          },
          chartObjects[0]!,
          { kind: 'rect', options: { name: 'mixed-order' } },
          { kind: 'image', source: jpegPath, options: { name: 'path-jpeg' } },
          {
            kind: 'image',
            source: 'https://example.test/source.gif',
            options: { name: 'url-gif' },
          },
          {
            kind: 'image',
            source: new Blob([sdkSvg(640, 360)], { type: 'image/svg+xml' }),
            options: {
              name: 'blob-svg',
              fallback: new Blob([sdkPngHeader(1, 1)], { type: 'image/png' }),
            },
          },
          {
            kind: 'image',
            source: Readable.from([png]),
            options: { name: 'stream-png' },
          },
          { kind: 'image', source: png, options: { name: 'bytes-png' } },
          ...chartObjects.slice(1),
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
            options: { name: 'combo' },
          },
        ],
      });

      expect(layout.background).toMatchObject({ kind: 'image', contentType: 'image/png' });
      expect(layout.shapes.slice(0, 4).map(({ kind }) => kind)).toEqual([
        'text', 'image', 'chart', 'shape',
      ]);
      const images = layout.shapes.filter(
        (shape): shape is ImageModel => shape instanceof ImageModel,
      );
      expect(images.map(({ name }) => name)).toEqual([
        'data-png', 'path-jpeg', 'url-gif', 'blob-svg', 'stream-png', 'bytes-png',
      ]);
      expect(images[0]?.sourceRectangle).toEqual({
        left: 21.875,
        top: 0,
        right: 21.875,
        bottom: 0,
      });
      expect(images[3]?.isSvg).toBe(true);
      expect(images[3]?.fallbackPartUri).toBeDefined();
      const charts = layout.shapes.filter(
        (shape): shape is ChartModel => shape instanceof ChartModel,
      );
      expect(charts).toHaveLength(10);
      expect(charts.slice(0, 9).map((chart) => chart.definition?.groups[0]?.type))
        .toEqual(chartTypes);
      expect(charts[9]?.definition?.groups.map(({ type, axis }) => [type, axis])).toEqual([
        ['bar', 'primary'],
        ['line', 'secondary'],
      ]);
      expect(new Set(charts.map(({ workbookPartUri }) => workbookPartUri)).size).toBe(10);
      const reopened = await PptxDocument.open(await document.write());
      const reopenedLayout = reopened.layouts.find(({ name }) => name === 'ASYNC')!;
      expect(reopenedLayout.background?.kind).toBe('image');
      expect(reopenedLayout.shapes.filter((shape) => shape instanceof ImageModel)).toHaveLength(6);
      expect(reopenedLayout.shapes.filter((shape) => shape instanceof ChartModel)).toHaveLength(10);
    } finally {
      fetchSpy.mockRestore();
      await rm(directory, { recursive: true, force: true });
    }

    const prepareFailure = PptxDocument.create();
    const prepareBefore = await sdkPackageSnapshot(prepareFailure);
    await expect(prepareFailure.defineSlideMaster({
      title: 'PREPARE-FAILURE',
      objects: [
        { kind: 'image', source: png },
        { kind: 'image', source: Uint8Array.of(1, 2, 3) },
      ],
    })).rejects.toThrow(/image|signature|svg/i);
    expect(await sdkPackageSnapshot(prepareFailure)).toEqual(prepareBefore);

    const commitFailure = PptxDocument.create();
    const commitBefore = await sdkPackageSnapshot(commitFailure);
    const originalSetPart = commitFailure.opcPackage.setPart.bind(commitFailure.opcPackage);
    const setPart = vi.spyOn(commitFailure.opcPackage, 'setPart')
      .mockImplementation((uri, bytes, contentType) => {
        if (uri.startsWith('/ppt/charts/')) throw new Error('async chart commit failed');
        return originalSetPart(uri, bytes, contentType);
      });
    await expect(commitFailure.defineSlideMaster({
      title: 'COMMIT-FAILURE',
      objects: [
        { kind: 'image', source: png },
        {
          kind: 'chart',
          groups: [{
            type: 'bar',
            series: [{ name: 'Revenue', categories: ['Q1'], values: [1] }],
          }],
        },
      ],
    })).rejects.toThrow('async chart commit failed');
    setPart.mockRestore();
    expect(await sdkPackageSnapshot(commitFailure)).toEqual(commitBefore);
  });

  it('replace delete slide master preserves identity, relinks, collects, and invalidates safely', async () => {
    const document = PptxDocument.create();
    document.masterLayoutTheme.copyMaster(document.masters[0]!.partUri);
    const secondMaster = document.masters[1]!;
    const original = await document.defineSlideMaster({
      title: 'LIFECYCLE-ORIGINAL',
      background: { kind: 'image-source', source: sdkPngHeader(2, 1) },
      margin: inches(0.25),
      slideNumber: { align: 'right' },
      objects: [
        { kind: 'text', text: 'Original text' },
        { kind: 'image', source: sdkPngHeader(3, 2), options: { name: 'Original image' } },
        {
          kind: 'chart',
          groups: [{
            type: 'bar',
            series: [{ name: 'Original chart', categories: ['Q1'], values: [1] }],
          }],
        },
      ],
    });
    const originalPartUri = original.partUri;
    const originalShape = original.shapes[0]!;
    const slide = document.addSlide({ masterName: 'LIFECYCLE-ORIGINAL' });
    const slideLayoutRelationship = slide.relationships.find(
      ({ type }) => type.endsWith('/slideLayout'),
    )!;
    const copiedRaw = document.masterLayoutTheme.copyLayout(original.partUri);
    const sharedCopy = document.layouts.find(({ partUri }) => partUri === copiedRaw.partUri)!;
    const oldTargets = document.opcPackage.relationships(original.partUri)
      .flatMap(({ resolvedTarget, type }) =>
        resolvedTarget && (type.endsWith('/image') || type.endsWith('/chart'))
          ? [resolvedTarget]
          : []);
    const oldWorkbookTargets = oldTargets.flatMap((target) =>
      document.opcPackage.relationships(target).flatMap(({ type, resolvedTarget }) =>
        type.endsWith('/package') && resolvedTarget ? [resolvedTarget] : []));
    expect(oldTargets.length).toBeGreaterThan(1);
    expect(oldWorkbookTargets).toHaveLength(1);

    const replacementDefinition = {
      title: 'LIFECYCLE-RENAMED',
      master: secondMaster,
      background: {
        kind: 'solid' as const,
        color: { kind: 'srgb' as const, value: '112233' },
      },
      margin: inches(0.5),
      objects: [{ kind: 'text' as const, text: 'Replacement text' }],
    };
    await document.replaceSlideMaster(original, replacementDefinition);
    expect(document.layouts.find(({ name }) => name === 'LIFECYCLE-RENAMED')).toBe(original);
    expect(original.masterPartUri).toBe(secondMaster.partUri);
    expect(original.background).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '112233' },
    });
    expect(original.margin).toEqual({
      top: inches(0.5),
      right: inches(0.5),
      bottom: inches(0.5),
      left: inches(0.5),
    });
    expect(original.slideNumber).toBeUndefined();
    expect(original.shapes.map(({ kind }) => kind)).toEqual(['text']);
    expect(slide.relationships.find(({ type }) => type.endsWith('/slideLayout'))).toEqual(
      slideLayoutRelationship,
    );
    expect(() => originalShape.name).toThrow(/stale/i);
    for (const target of oldTargets) expect(document.opcPackage.hasPart(target)).toBe(true);
    for (const target of oldWorkbookTargets) expect(document.opcPackage.hasPart(target)).toBe(true);

    document.title = 'Unrelated package mutation';
    const journal = [...document.opcPackage.mutations];
    await document.replaceSlideMaster(original, replacementDefinition);
    expect(document.opcPackage.mutations).toEqual(journal);
    original.addText('Direct drift');
    await document.replaceSlideMaster(original, replacementDefinition);
    expect(original.shapes.map((shape) => shape instanceof ShapeModel ? shape.text : ''))
      .toEqual(['Replacement text']);

    document.deleteSlideMaster(sharedCopy);
    expect(oldTargets.filter((target) => document.opcPackage.hasPart(target))).toEqual([]);
    expect(oldWorkbookTargets.filter((target) => document.opcPackage.hasPart(target))).toEqual([]);
    const { output: _beforeRejectedOutput, ...beforeRejectedDelete } =
      await sdkPackageSnapshot(document) as {
        readonly output: Uint8Array;
        readonly [key: string]: unknown;
      };
    expect(() => document.deleteSlideMaster(original)).toThrow(/used|slide/i);
    const { output: _afterRejectedOutput, ...afterRejectedDelete } =
      await sdkPackageSnapshot(document) as {
        readonly output: Uint8Array;
        readonly [key: string]: unknown;
      };
    expect(afterRejectedDelete).toEqual(beforeRejectedDelete);

    const replacement = await document.defineSlideMaster({
      title: 'LIFECYCLE-REPLACEMENT',
      master: secondMaster,
    });
    document.deleteSlideMaster(original, replacement);
    expect(slide.relationships.find(({ type }) => type.endsWith('/slideLayout'))).toMatchObject({
      id: slideLayoutRelationship.id,
      resolvedTarget: replacement.partUri,
    });
    expect(document.layouts).not.toContain(original);
    expect(() => original.name).toThrow(/stale|deleted/i);

    const reused = await document.defineSlideMaster({ title: 'LIFECYCLE-REUSED' });
    expect(reused.partUri).toBe(originalPartUri);
    expect(reused).not.toBe(original);
    expect(reused.name).toBe('LIFECYCLE-REUSED');
    expect(reused.margin).toBeUndefined();

    const rawDeleted = await document.defineSlideMaster({ title: 'RAW-DELETE' });
    const rawDeletedPartUri = rawDeleted.partUri;
    document.masterLayoutTheme.deleteLayout(rawDeletedPartUri);
    const rawReused = await document.defineSlideMaster({ title: 'RAW-REUSED' });
    expect(rawReused.partUri).toBe(rawDeletedPartUri);
    expect(rawReused).not.toBe(rawDeleted);
    expect(() => rawDeleted.name).toThrow(/stale|deleted/i);

    const exclusiveDefinition = {
      title: 'EXCLUSIVE-REUSE',
      objects: [
        { kind: 'image' as const, source: sdkPngHeader(4, 3) },
        {
          kind: 'chart' as const,
          groups: [{
            type: 'bar' as const,
            series: [{ name: 'Exclusive', categories: ['Q1'], values: [7] }],
          }],
        },
      ],
    };
    const exclusive = await document.defineSlideMaster(exclusiveDefinition);
    const exclusiveTargets = document.opcPackage.relationships(exclusive.partUri)
      .flatMap(({ type, resolvedTarget }) =>
        resolvedTarget && (type.endsWith('/image') || type.endsWith('/chart'))
          ? [resolvedTarget]
          : []);
    const exclusiveImageUri = exclusiveTargets.find((uri) => uri.includes('/media/'))!;
    const exclusiveChartUri = exclusiveTargets.find((uri) => uri.includes('/charts/'))!;
    const exclusiveWorkbookUri = document.opcPackage.relationships(exclusiveChartUri)[0]!
      .resolvedTarget!;
    const exclusiveImageBytes = document.opcPackage.requirePart(exclusiveImageUri).bytes;
    document.opcPackage.setPart(exclusiveImageUri, Uint8Array.of(1, 2, 3), 'image/png');
    await document.replaceSlideMaster(exclusive, exclusiveDefinition);
    expect(document.opcPackage.relationships(exclusive.partUri)
      .flatMap(({ type, resolvedTarget }) =>
        resolvedTarget && (type.endsWith('/image') || type.endsWith('/chart'))
          ? [resolvedTarget]
          : []))
      .toEqual(exclusiveTargets);
    expect(document.opcPackage.requirePart(exclusiveImageUri).bytes).toEqual(exclusiveImageBytes);
    expect(document.opcPackage.relationships(exclusiveChartUri)[0]?.resolvedTarget)
      .toBe(exclusiveWorkbookUri);
  });

  it('replace delete slide master rolls back and rejects creating slides without a layout', async () => {
    const document = PptxDocument.create();
    const layout = await document.defineSlideMaster({
      title: 'ROLLBACK-ORIGINAL',
      objects: [{ kind: 'text', text: 'Rollback original' }],
    });
    const rollbackOriginalShape = layout.shapes[0] as ShapeModel;
    const { output: _beforeOutput, ...before } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    const originalSetPart = document.opcPackage.setPart.bind(document.opcPackage);
    const setPart = vi.spyOn(document.opcPackage, 'setPart')
      .mockImplementation((uri, bytes, contentType) => {
        if (uri.startsWith('/ppt/charts/')) throw new Error('replace chart failed');
        return originalSetPart(uri, bytes, contentType);
      });
    await expect(document.replaceSlideMaster(layout, {
      title: 'ROLLBACK-RENAMED',
      margin: inches(0.5),
      objects: [
        { kind: 'image', source: sdkPngHeader(1, 1) },
        {
          kind: 'chart',
          groups: [{
            type: 'bar',
            series: [{ name: 'Rollback', categories: ['Q1'], values: [1] }],
          }],
        },
      ],
    })).rejects.toThrow('replace chart failed');
    setPart.mockRestore();
    const { output: _afterOutput, ...after } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(after).toEqual(before);
    expect(layout.name).toBe('ROLLBACK-ORIGINAL');
    expect(layout.margin).toBeUndefined();
    expect(layout.shapes).toEqual([rollbackOriginalShape]);
    expect(rollbackOriginalShape.text).toBe('Rollback original');

    const concurrent = PptxDocument.create();
    const concurrentTarget = await concurrent.defineSlideMaster({ title: 'CONCURRENT-TARGET' });
    let signalRead!: () => void;
    let resumeRead!: () => void;
    const readStarted = new Promise<void>((resolve) => {
      signalRead = resolve;
    });
    const readPaused = new Promise<void>((resolve) => {
      resumeRead = resolve;
    });
    const pendingReplace = concurrent.replaceSlideMaster(concurrentTarget, {
      title: 'CONCURRENT-OLD',
      objects: [{
        kind: 'image',
        source: {
          async *[Symbol.asyncIterator]() {
            signalRead();
            await readPaused;
            yield sdkPngHeader(2, 2);
          },
        },
      }],
    });
    await readStarted;
    const concurrentPartUri = concurrentTarget.partUri;
    concurrent.deleteSlideMaster(concurrentTarget);
    const concurrentReused = await concurrent.defineSlideMaster({ title: 'CONCURRENT-REUSED' });
    expect(concurrentReused.partUri).toBe(concurrentPartUri);
    resumeRead();
    await expect(pendingReplace).rejects.toThrow(/detached|stale/i);
    expect(concurrentReused.name).toBe('CONCURRENT-REUSED');

    const withFallback = PptxDocument.create();
    const formerDefault = withFallback.layouts[0]!;
    const fallback = await withFallback.defineSlideMaster({ title: 'FALLBACK' });
    withFallback.deleteSlideMaster(formerDefault);
    const fallbackSlide = withFallback.addSlide();
    expect(fallbackSlide.relationships.find(({ type }) => type.endsWith('/slideLayout'))
      ?.resolvedTarget).toBe(fallback.partUri);

    const empty = PptxDocument.create();
    const onlyLayout = empty.layouts[0]!;
    empty.deleteSlideMaster(onlyLayout);
    const emptyBefore = await sdkPackageSnapshot(empty);
    expect(() => empty.addSlide()).toThrow(/layout/i);
    expect(await sdkPackageSnapshot(empty)).toEqual(emptyBefore);
  });

  it('edits and reopens direct layout master backgrounds', async () => {
    const document = PptxDocument.create();
    const layout = document.layouts[0]!;
    const master = document.masters[0]!;
    layout.background = {
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 20,
    };
    master.background = { kind: 'none' };
    expect(layout.background).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 20,
    });
    expect(master.background).toEqual({ kind: 'none' });

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.layouts[0]?.background).toEqual(layout.background);
    expect(reopened.masters[0]?.background).toEqual(master.background);
    reopened.layouts[0]!.background = undefined;
    reopened.masters[0]!.background = undefined;
    expect(reopened.layouts[0]?.background).toBeUndefined();
    expect(reopened.masters[0]?.background).toBeUndefined();
  });

  it('exposes stable semantic master layout models with editable content', async () => {
    const document = PptxDocument.create();
    const layout = document.layouts[0]!;
    const master = document.masters[0]!;
    expect(layout).toBeInstanceOf(SlideLayoutModel);
    expect(master).toBeInstanceOf(SlideMasterModel);
    expect(document.layouts[0]).toBe(layout);
    expect(document.masters[0]).toBe(master);
    expect(master.layouts[0]).toBe(layout);
    expect(document.masterLayoutTheme.layouts[0]?.partUri).toBe(layout.partUri);
    expect(document.masterLayoutTheme.masters[0]?.partUri).toBe(master.partUri);

    const text = layout.addText('Inherited layout text', {
      x: inches(1),
      y: inches(2),
      width: inches(3),
      height: inches(4),
    });
    const rich = layout.addRichText([{ runs: [{ text: 'Rich layout text' }] }]);
    const shape = master.addShape('rect', {
      x: inches(0.1),
      y: inches(0.2),
      width: inches(0.3),
      height: inches(0.4),
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
    });
    const image = layout.addImage(sdkPngHeader(2, 3), {
      contentType: 'image/png',
      x: inches(0.5),
      y: inches(0.6),
      width: inches(0.7),
      height: inches(0.8),
    });
    const chart = await master.addChart('bar', [{
      name: 'Revenue',
      categories: ['Q1', 'Q2'],
      values: [10, 20],
    }]);

    expect(layout.shapes.find(({ id }) => id === text.id)).toBe(text);
    expect(layout.shapes.find(({ id }) => id === rich.id)).toBe(rich);
    expect(layout.shapes.find(({ id }) => id === image.id)).toBe(image);
    expect(master.shapes.find(({ id }) => id === shape.id)).toBe(shape);
    expect(master.shapes.find(({ id }) => id === chart.id)).toBe(chart);
    text.text = 'Edited layout text';
    shape.setTransform({ x: inches(0.9) });
    expect(layout.shapes.find(({ id }) => id === text.id)).toBe(text);
    expect(text.text).toBe('Edited layout text');
    expect(shape.transform.x).toBe(inches(0.9));

    layout.slideNumber = { x: 200, align: 'center' };
    master.slideNumber = { x: 300, align: 'right' };
    expect(layout.slideNumber).toMatchObject({ x: 200, align: 'center' });
    expect(master.slideNumber).toMatchObject({ x: 300, align: 'right' });

    const copiedRaw = document.masterLayoutTheme.copyLayout(layout.partUri);
    const copied = document.layouts.find(({ partUri }) => partUri === copiedRaw.partUri)!;
    expect(document.layouts.find(({ partUri }) => partUri === copiedRaw.partUri)).toBe(copied);
    document.masterLayoutTheme.deleteLayout(copied.partUri);
    expect(document.layouts).not.toContain(copied);
    expect(() => copied.name).toThrow();

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.layouts[0]).toBe(reopened.layouts[0]);
    expect(reopened.masters[0]?.layouts[0]).toBe(reopened.layouts[0]);
    expect(reopened.layouts[0]?.shapes.map(({ kind }) => kind)).toEqual([
      'text',
      'text',
      'image',
      'text',
    ]);
    expect(reopened.masters[0]?.shapes.map(({ kind }) => kind)).toEqual([
      'shape',
      'chart',
      'text',
    ]);
  });

  it('selects named slide layouts with sections in all six presentation formats', async () => {
    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const document = PptxDocument.create({ format });
      const rawMaster = document.masterLayoutTheme.masters[0]!;
      const defaultLayout = document.masterLayoutTheme.layouts[0]!;
      const defaultXml = new TextDecoder().decode(
        document.opcPackage.requirePart(defaultLayout.partUri).bytes,
      );
      const brand = document.masterLayoutTheme.createLayout(
        rawMaster.partUri,
        defaultXml.replace('name="DEFAULT"', `name="BRAND-${format}"`),
      );
      const section = document.addSection({ title: `Named ${format}` });

      const slide = document.addSlide({
        masterName: `BRAND-${format}`,
        sectionTitle: section.title,
      });
      const inherited = document.addSlide({ sectionTitle: section.title });
      expect(slide.relationships.find(
        ({ type }) => type.endsWith('/slideLayout'),
      )?.resolvedTarget).toBe(brand.partUri);
      expect(inherited.relationships.find(
        ({ type }) => type.endsWith('/slideLayout'),
      )?.resolvedTarget).toBe(brand.partUri);
      expect(document.sections?.[0]?.slideIds).toEqual([slide.slideId, inherited.slideId]);

      const reopened = await PptxDocument.open(await document.write());
      const reopenedSlide = reopened.slides[0]!;
      expect(reopened.format).toBe(format);
      expect(reopenedSlide.relationships.find(
        ({ type }) => type.endsWith('/slideLayout'),
      )?.resolvedTarget).toBe(brand.partUri);
      expect(reopened.masterLayoutTheme.layouts.find(
        ({ partUri }) => partUri === brand.partUri,
      )?.name).toBe(`BRAND-${format}`);
      expect(reopened.slides[1]?.relationships.find(
        ({ type }) => type.endsWith('/slideLayout'),
      )?.resolvedTarget).toBe(brand.partUri);
      expect(reopened.sections?.[0]?.slideIds).toEqual(
        reopened.slides.map(({ slideId }) => slideId),
      );
    }
  });

  it('rejects unsafe named slide layout selection without mutation', async () => {
    const document = PptxDocument.create();
    const rawMaster = document.masterLayoutTheme.masters[0]!;
    const defaultLayout = document.masterLayoutTheme.layouts[0]!;
    const defaultXml = new TextDecoder().decode(
      document.opcPackage.requirePart(defaultLayout.partUri).bytes,
    );
    document.masterLayoutTheme.createLayout(
      rawMaster.partUri,
      defaultXml.replace('name="DEFAULT"', 'name="DUPLICATE"'),
    );
    document.masterLayoutTheme.createLayout(
      rawMaster.partUri,
      defaultXml.replace('name="DEFAULT"', 'name="DUPLICATE"'),
    );
    let accessorCalls = 0;
    const accessor = Object.defineProperty({}, 'masterName', {
      enumerable: true,
      get: () => {
        accessorCalls += 1;
        return 'DEFAULT';
      },
    });
    const before = await sdkPackageSnapshot(document);

    expect(() => document.addSlide({ masterName: 'MISSING' })).toThrow(RangeError);
    expect(() => document.addSlide({ masterName: 'DUPLICATE' })).toThrow(ModelParseError);
    for (const options of [
      { masterName: '   ' },
      { masterName: 'BAD\u0000NAME' },
      { masterName: 'DEFAULT', extra: true },
      { masterName: 'DEFAULT', [Symbol('extra')]: true },
      accessor,
    ]) {
      expect(() => document.addSlide(options as never)).toThrow(TypeError);
    }
    expect(accessorCalls).toBe(0);
    expect(await sdkPackageSnapshot(document)).toEqual(before);
  });

  it('creates slide numbers from strict public options and rejects invalid starts', () => {
    const createOptions: CreatePresentationOptions = { firstSlideNumber: 0 };
    const numberOptions: SlideNumberOptions = {
      x: 0,
      y: 0,
      width: 800_000,
      height: 300_000,
      align: 'justify',
      rtl: true,
      valign: 'middle',
      margin: [1, 2, 3, 4],
      style: {
        fontFamily: 'Aptos',
        fontSize: 18,
        lang: 'zh-CN',
        bold: true,
        italic: true,
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 20,
      },
    };
    const margin: SlideNumberMarginInput = 0;
    const margins: SlideNumberMargins = { left: 1 };
    const color: SlideNumberColor = { kind: 'srgb', value: 'FF3399' };
    const styleOptions: SlideNumberTextStyleOptions = { color };
    const style: SlideNumberTextStyle = {
      lang: 'en-US', bold: false, italic: false, ...styleOptions,
    };
    const document = PptxDocument.create(createOptions);
    const control = PptxDocument.create();
    const slide = document.addSlide();
    slide.slideNumber = { ...numberOptions, margin };
    const value: Readonly<SlideNumber> | undefined = slide.slideNumber;
    expect(document.firstSlideNumber).toBe(0);
    expect(value).toMatchObject({ align: 'justify', rtl: true });
    expect(slide.title.text).toBe('');
    expect([margins, style]).toHaveLength(2);
    const presentationXml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.presentationPartUri).bytes,
    );
    expect(presentationXml).toContain(' firstSlideNum="0"');
    expect(presentationXml.match(/\bfirstSlideNum=/g)).toHaveLength(1);
    expect(document.opcPackage.mutations.slice(0, control.opcPackage.mutations.length))
      .toEqual(control.opcPackage.mutations);

    const createSpy = vi.spyOn(OpcPackage, 'create');
    try {
      for (const invalid of [1.5, Number.NaN, 2_147_483_648, -2_147_483_649]) {
        expect(() => PptxDocument.create({ firstSlideNumber: invalid })).toThrow();
      }
      expect(createSpy).not.toHaveBeenCalled();
    } finally {
      createSpy.mockRestore();
    }
  });

  it('round-trips all three slide-number owners twice in all six formats', async () => {
    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({ format, firstSlideNumber: 5 });
      const firstSlide = created.addSlide();
      const secondSlide = created.addSlide();
      firstSlide.slideNumber = { align: 'left' };
      secondSlide.slideNumber = {
        align: 'justify',
        rtl: true,
        style: {
          italic: true,
          color: { kind: 'srgb', value: 'FF3399' },
          transparency: 25,
        },
      };
      created.layouts[0]!.slideNumber = { x: 200, align: 'center' };
      created.masters[0]!.slideNumber = { x: 300, align: 'right' };
      const presentationContentType = created.opcPackage
        .requirePart(created.presentationPartUri).contentType;

      const first = await PptxDocument.open(await created.write());
      expect(first.slides[0]?.slideNumber?.align).toBe('left');
      expect(first.slides[1]?.slideNumber?.align).toBe('justify');
      expect(first.layouts[0]?.slideNumber?.x).toBe(200);
      expect(first.masters[0]?.slideNumber?.x).toBe(300);
      first.layouts[0]!.slideNumber = undefined;
      first.masters[0]!.slideNumber = undefined;
      expect(first.layouts[0]?.slideNumber).toBeUndefined();
      expect(first.masters[0]?.slideNumber).toBeUndefined();
      first.layouts[0]!.slideNumber = { x: 250, align: 'center' };
      first.masters[0]!.slideNumber = { x: 350, align: 'right' };
      const duplicate = first.duplicateSlide(0);
      first.moveSlide(first.slides.indexOf(duplicate), 0);
      first.deleteSlide(first.slides.indexOf(first.slides.find(
        ({ partUri }) => partUri === firstSlide.partUri,
      )!));

      const second = await PptxDocument.open(await first.write());
      expect(second.format).toBe(format);
      expect(second.formatProfile).toEqual(PRESENTATION_FORMAT_PROFILES[format]);
      expect(second.opcPackage.requirePart(second.presentationPartUri).contentType)
        .toBe(presentationContentType);
      expect(second.firstSlideNumber).toBe(5);
      expect(second.slides).toHaveLength(2);
      expect(second.slides.map((slide) => sdkSlideNumberCache(second, slide.partUri)))
        .toEqual(['5', '6']);
      expect(second.layouts[0]?.slideNumber).toMatchObject({ x: 250, align: 'center' });
      expect(second.masters[0]?.slideNumber).toMatchObject({ x: 350, align: 'right' });
      expect(sdkSlideNumberCache(second, second.layouts[0]!.partUri)).toBe('‹#›');
      expect(sdkSlideNumberCache(second, second.masters[0]!.partUri)).toBe('‹#›');
      expect(new TextDecoder().decode(
        second.opcPackage.requirePart(second.masters[0]!.partUri).bytes,
      )).toContain('sldNum="1"');
      expect(validatePackage(second.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);
    }
  });

  it('round-trips materialized slide default colors twice in all six formats', async () => {
    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({ format });
      const source = created.addSlide();
      source.color = { kind: 'scheme', value: 'accent2' };
      source.addText('Plain inherited');
      source.addRichText([{
        runs: [
          { text: 'Rich inherited' },
          { text: 'Rich override', style: { color: { kind: 'srgb', value: '00AA00' } } },
        ],
      }]);
      const duplicate = created.duplicateSlide(0);
      expect(duplicate.color).toBe(source.color);
      duplicate.addText('Duplicate inherited');

      const first = await PptxDocument.open(await created.write());
      expect(first.format).toBe(format);
      expect(first.slides.map(({ color }) => color)).toEqual([undefined, undefined]);
      const firstColors = first.slides.map((slide) => slide.shapes
        .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
        .map(({ richText }) => richText.flatMap(({ runs }) =>
          runs.map(({ style }) => style?.color))));
      expect(firstColors).toEqual([
        [
          [{ kind: 'scheme', value: 'accent2' }],
          [
            { kind: 'scheme', value: 'accent2' },
            { kind: 'srgb', value: '00AA00' },
          ],
        ],
        [
          [{ kind: 'scheme', value: 'accent2' }],
          [
            { kind: 'scheme', value: 'accent2' },
            { kind: 'srgb', value: '00AA00' },
          ],
          [{ kind: 'scheme', value: 'accent2' }],
        ],
      ]);

      const second = await PptxDocument.open(await first.write());
      expect(second.format).toBe(format);
      expect(second.formatProfile).toEqual(PRESENTATION_FORMAT_PROFILES[format]);
      expect(second.slides.map(({ color }) => color)).toEqual([undefined, undefined]);
      expect(validatePackage(second.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);
    }
  });

  it('creates text fills across slide layout master and placeholder owners', async () => {
    const document = PptxDocument.create();
    const layout = document.layouts[0]!;
    const master = document.masters[0]!;
    const layoutText = layout.addText('Layout fill', {
      name: 'layout_fill',
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
    });
    const masterText = master.addRichText([{ runs: [{ text: 'Master fill' }] }], {
      name: 'master_fill',
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: '112233' },
        transparency: 40,
      },
    });
    const layoutPlaceholder = layout.addPlaceholder('Layout prompt', {
      name: 'title_fill',
      type: 'title',
      index: 190,
      x: inches(1),
      y: inches(1),
      width: inches(8),
      height: inches(1),
      fill: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 10,
      },
    });

    expect(layout.shapes.find(({ id }) => id === layoutText.id)).toBe(layoutText);
    expect(master.shapes.find(({ id }) => id === masterText.id)).toBe(masterText);
    expect(layout.placeholders.find(({ id }) => id === layoutPlaceholder.id))
      .toBe(layoutPlaceholder);
    expect(layoutText.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
    });
    expect(masterText.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '112233' },
      transparency: 40,
    });
    expect(layoutPlaceholder.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 10,
    });

    const slide = document.addSlide({ masterName: layout.name });
    const materialized = slide.placeholders.find(({ name }) => name === 'title_fill')!;
    const materializedState = {
      id: materialized.id,
      name: materialized.name,
      transform: materialized.transform,
      placeholder: materialized.placeholder,
    };
    const populated = slide.addText('Populated fill', {
      placeholder: 'title_fill',
      fill: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent3' },
        transparency: 25,
      },
    });
    expect(populated).toBe(slide.shapes.find(({ id }) => id === materializedState.id));
    expect({
      id: populated.id,
      name: populated.name,
      transform: populated.transform,
      placeholder: populated.placeholder,
    }).toEqual(materializedState);
    expect(populated.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent3' },
      transparency: 25,
    });
    expect(layoutPlaceholder.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 10,
    });
    expect(materialized).not.toBe(populated);

    const declarative = await document.defineSlideMaster({
      title: 'TEXT-FILLS',
      objects: [
        {
          kind: 'text',
          text: 'Declarative text fill',
          options: {
            name: 'declarative_text_fill',
            fill: {
              kind: 'solid',
              color: { kind: 'scheme', value: 'accent4' },
              transparency: 0,
            },
          },
        },
        {
          kind: 'placeholder',
          text: 'Declarative prompt',
          options: {
            name: 'declarative_title_fill',
            type: 'title',
            index: 191,
            fill: {
              kind: 'solid',
              color: { kind: 'srgb', value: 'AABBCC' },
              transparency: 50,
            },
          },
        },
      ],
    });
    const declarativeText = declarative.shapes.find(
      ({ name }) => name === 'declarative_text_fill',
    ) as ShapeModel;
    const declarativePlaceholder = declarative.placeholders.find(
      ({ name }) => name === 'declarative_title_fill',
    )!;
    expect(declarativeText.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent4' },
      transparency: 0,
    });
    expect(declarativePlaceholder.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'AABBCC' },
      transparency: 50,
    });
    const declarativeSlide = document.addSlide({ masterName: declarative.name });
    const declarativePopulated = declarativeSlide.addText('Declarative populated', {
      placeholder: 'declarative_title_fill',
      fill: { kind: 'none' },
    });
    expect(declarativePopulated.fill).toEqual({ kind: 'none' });
    expect(declarativePlaceholder.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'AABBCC' },
      transparency: 50,
    });

    const duplicate = document.duplicateSlide(document.slides.indexOf(slide));
    const duplicatePopulated = duplicate.shapes.find(
      ({ name }) => name === populated.name,
    ) as ShapeModel;
    duplicatePopulated.fill = { kind: 'none' };
    expect(duplicatePopulated.fill).toEqual({ kind: 'none' });
    expect(populated.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent3' },
      transparency: 25,
    });

    const beforeRollbackShapes = [...slide.shapes];
    let rolledBack: ShapeModel | undefined;
    expect(() => document.transaction(() => {
      rolledBack = slide.addText('Rolled back fill', {
        fill: { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } },
      });
      throw new Error('restore created text fill');
    })).toThrow('restore created text fill');
    expect(slide.shapes).toEqual(beforeRollbackShapes);
    expect(slide.shapes.find(({ id }) => id === populated.id)).toBe(populated);
    expect(() => rolledBack!.fill).toThrow(ModelParseError);

    let signalRead!: () => void;
    let resumeRead!: () => void;
    const readStarted = new Promise<void>((resolve) => { signalRead = resolve; });
    const readPaused = new Promise<void>((resolve) => { resumeRead = resolve; });
    const detachedFill: {
      kind: 'solid';
      color: { kind: 'srgb'; value: string };
      transparency: number;
    } = {
      kind: 'solid',
      color: { kind: 'srgb', value: '123456' },
      transparency: 30,
    };
    const pendingDetached = document.defineSlideMaster({
      title: 'DETACHED-TEXT-FILL',
      objects: [
        {
          kind: 'text',
          text: 'Detached text fill',
          options: { name: 'detached_text_fill', fill: detachedFill },
        },
        {
          kind: 'image',
          source: {
            async *[Symbol.asyncIterator]() {
              signalRead();
              await readPaused;
              yield sdkPngHeader(1, 1);
            },
          },
        },
      ],
    });
    await readStarted;
    detachedFill.color.value = 'FFFFFF';
    detachedFill.transparency = 90;
    resumeRead();
    const detachedLayout = await pendingDetached;
    expect((detachedLayout.shapes.find(
      ({ name }) => name === 'detached_text_fill',
    ) as ShapeModel).fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '123456' },
      transparency: 30,
    });

    const reopened = await PptxDocument.open(await document.write());
    const second = await PptxDocument.open(await reopened.write());
    expect((second.layouts.find(({ name }) => name === layout.name)!.shapes.find(
      ({ name }) => name === 'layout_fill',
    ) as ShapeModel).fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
    });
    expect((second.masters[0]!.shapes.find(
      ({ name }) => name === 'master_fill',
    ) as ShapeModel).fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '112233' },
      transparency: 40,
    });
    expect((second.slides.find(({ partUri }) => partUri === slide.partUri)!.shapes.find(
      ({ name }) => name === populated.name,
    ) as ShapeModel).fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent3' },
      transparency: 25,
    });
    expect(validatePackage(second.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      const formattedLayout = formatted.layouts[0]!;
      formattedLayout.addText('Formatted layout fill', {
        name: 'formatted_layout_fill',
        fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent5' } },
      });
      formatted.masters[0]!.addText('Formatted master fill', {
        name: 'formatted_master_fill',
        fill: {
          kind: 'solid',
          color: { kind: 'srgb', value: '445566' },
          transparency: 100,
        },
      });
      formattedLayout.addPlaceholder('Formatted prompt', {
        name: 'formatted_title_fill',
        type: 'title',
        index: 192,
        fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent6' } },
      });
      const formattedSlide = formatted.addSlide({ masterName: formattedLayout.name });
      formattedSlide.addText('Formatted populated fill', {
        placeholder: 'formatted_title_fill',
        fill: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 25,
        },
      });

      const first = await PptxDocument.open(await formatted.write());
      const formattedSecond = await PptxDocument.open(await first.write());
      expect(formattedSecond.format).toBe(format);
      expect(formattedSecond.formatProfile).toEqual(PRESENTATION_FORMAT_PROFILES[format]);
      expect((formattedSecond.layouts[0]!.shapes.find(
        ({ name }) => name === 'formatted_layout_fill',
      ) as ShapeModel).fill).toEqual({
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent5' },
      });
      expect((formattedSecond.masters[0]!.shapes.find(
        ({ name }) => name === 'formatted_master_fill',
      ) as ShapeModel).fill).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: '445566' },
        transparency: 100,
      });
      expect((formattedSecond.slides[0]!.shapes.find(
        ({ name }) => name === 'formatted_title_fill',
      ) as ShapeModel).fill).toEqual({
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 25,
      });
      expect(validatePackage(formattedSecond.opcPackage).filter(
        ({ severity }) => severity === 'error',
      )).toEqual([]);
    }
  });

  it('creates text lines across slide layout master and placeholder owners', async () => {
    const document = PptxDocument.create();
    const layout = document.layouts[0]!;
    const master = document.masters[0]!;
    const layoutText = layout.addText('Layout line', {
      name: 'layout_line',
      line: { kind: 'line', color: { kind: 'scheme', value: 'accent1' } },
    });
    const masterText = master.addRichText([{ runs: [{ text: 'Master line' }] }], {
      name: 'master_line',
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
        transparency: 40,
        width: 2,
        dash: 'dash',
      },
    });
    const layoutPlaceholder = layout.addPlaceholder('Layout prompt', {
      name: 'title_line',
      type: 'title',
      index: 193,
      x: inches(1),
      y: inches(1),
      width: inches(8),
      height: inches(1),
      line: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 10,
        width: 0,
        dash: 'sysDot',
      },
    });

    expect(layout.shapes.find(({ id }) => id === layoutText.id)).toBe(layoutText);
    expect(master.shapes.find(({ id }) => id === masterText.id)).toBe(masterText);
    expect(layout.placeholders.find(({ id }) => id === layoutPlaceholder.id))
      .toBe(layoutPlaceholder);
    expect(layoutText.line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent1' },
      width: 1,
      dash: 'solid',
    });
    expect(masterText.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      transparency: 40,
      width: 2,
      dash: 'dash',
    });
    expect(layoutPlaceholder.line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 10,
      width: 0,
      dash: 'sysDot',
    });

    const slide = document.addSlide({ masterName: layout.name });
    const materialized = slide.placeholders.find(({ name }) => name === 'title_line')!;
    const materializedState = {
      id: materialized.id,
      name: materialized.name,
      transform: materialized.transform,
      placeholder: materialized.placeholder,
    };
    const populated = slide.addText('Populated line', {
      placeholder: 'title_line',
      line: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent3' },
        transparency: 25,
        width: 2.5,
        dash: 'dashDot',
      },
    });
    expect(populated).toBe(slide.shapes.find(({ id }) => id === materializedState.id));
    expect({
      id: populated.id,
      name: populated.name,
      transform: populated.transform,
      placeholder: populated.placeholder,
    }).toEqual(materializedState);
    expect(populated.line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent3' },
      transparency: 25,
      width: 2.5,
      dash: 'dashDot',
    });
    expect(layoutPlaceholder.line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 10,
      width: 0,
      dash: 'sysDot',
    });
    expect(materialized).not.toBe(populated);

    const declarative = await document.defineSlideMaster({
      title: 'TEXT-LINES',
      objects: [
        {
          kind: 'text',
          text: 'Declarative text line',
          options: {
            name: 'declarative_text_line',
            line: {
              kind: 'line',
              color: { kind: 'scheme', value: 'accent4' },
              transparency: 0,
            },
          },
        },
        {
          kind: 'placeholder',
          text: 'Declarative prompt',
          options: {
            name: 'declarative_title_line',
            type: 'title',
            index: 194,
            line: {
              kind: 'line',
              color: { kind: 'srgb', value: 'AABBCC' },
              transparency: 50,
              width: 3,
              dash: 'sysDash',
            },
          },
        },
      ],
    });
    const declarativeText = declarative.shapes.find(
      ({ name }) => name === 'declarative_text_line',
    ) as ShapeModel;
    const declarativePlaceholder = declarative.placeholders.find(
      ({ name }) => name === 'declarative_title_line',
    )!;
    expect(declarativeText.line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent4' },
      transparency: 0,
      width: 1,
      dash: 'solid',
    });
    expect(declarativePlaceholder.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: 'AABBCC' },
      transparency: 50,
      width: 3,
      dash: 'sysDash',
    });
    const declarativeSlide = document.addSlide({ masterName: declarative.name });
    const declarativePopulated = declarativeSlide.addText('Declarative populated', {
      placeholder: 'declarative_title_line',
      line: { kind: 'none' },
    });
    expect(declarativePopulated.line).toEqual({ kind: 'none' });
    expect(declarativePlaceholder.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: 'AABBCC' },
      transparency: 50,
      width: 3,
      dash: 'sysDash',
    });

    const duplicate = document.duplicateSlide(document.slides.indexOf(slide));
    const duplicatePopulated = duplicate.shapes.find(
      ({ name }) => name === populated.name,
    ) as ShapeModel;
    duplicatePopulated.line = { kind: 'none' };
    expect(duplicatePopulated.line).toEqual({ kind: 'none' });
    expect(populated.line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent3' },
      transparency: 25,
      width: 2.5,
      dash: 'dashDot',
    });

    const beforeRollbackShapes = [...slide.shapes];
    let rolledBack: ShapeModel | undefined;
    expect(() => document.transaction(() => {
      rolledBack = slide.addText('Rolled back line', {
        line: { kind: 'line', color: { kind: 'srgb', value: 'FF0000' } },
      });
      throw new Error('restore created text line');
    })).toThrow('restore created text line');
    expect(slide.shapes).toEqual(beforeRollbackShapes);
    expect(slide.shapes.find(({ id }) => id === populated.id)).toBe(populated);
    expect(() => rolledBack!.line).toThrow(ModelParseError);

    let signalRead!: () => void;
    let resumeRead!: () => void;
    const readStarted = new Promise<void>((resolve) => { signalRead = resolve; });
    const readPaused = new Promise<void>((resolve) => { resumeRead = resolve; });
    const detachedLine: {
      kind: 'line';
      color: { kind: 'srgb'; value: string };
      transparency: number;
      width: number;
      dash: ShapeLineDash;
    } = {
      kind: 'line',
      color: { kind: 'srgb', value: '123456' },
      transparency: 30,
      width: 1.5,
      dash: 'lgDashDot',
    };
    const pendingDetached = document.defineSlideMaster({
      title: 'DETACHED-TEXT-LINE',
      objects: [
        {
          kind: 'text',
          text: 'Detached text line',
          options: { name: 'detached_text_line', line: detachedLine },
        },
        {
          kind: 'image',
          source: {
            async *[Symbol.asyncIterator]() {
              signalRead();
              await readPaused;
              yield sdkPngHeader(1, 1);
            },
          },
        },
      ],
    });
    await readStarted;
    detachedLine.color.value = 'FFFFFF';
    detachedLine.transparency = 90;
    detachedLine.width = 9;
    detachedLine.dash = 'solid';
    resumeRead();
    const detachedLayout = await pendingDetached;
    expect((detachedLayout.shapes.find(
      ({ name }) => name === 'detached_text_line',
    ) as ShapeModel).line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '123456' },
      transparency: 30,
      width: 1.5,
      dash: 'lgDashDot',
    });

    const reopened = await PptxDocument.open(await document.write());
    const second = await PptxDocument.open(await reopened.write());
    expect((second.layouts.find(({ name }) => name === layout.name)!.shapes.find(
      ({ name }) => name === 'layout_line',
    ) as ShapeModel).line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent1' },
      width: 1,
      dash: 'solid',
    });
    expect((second.masters[0]!.shapes.find(
      ({ name }) => name === 'master_line',
    ) as ShapeModel).line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      transparency: 40,
      width: 2,
      dash: 'dash',
    });
    expect((second.slides.find(({ partUri }) => partUri === slide.partUri)!.shapes.find(
      ({ name }) => name === populated.name,
    ) as ShapeModel).line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent3' },
      transparency: 25,
      width: 2.5,
      dash: 'dashDot',
    });
    expect(validatePackage(second.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      const formattedLayout = formatted.layouts[0]!;
      formattedLayout.addText('Formatted layout line', {
        name: 'formatted_layout_line',
        line: { kind: 'line', color: { kind: 'scheme', value: 'accent5' } },
      });
      formatted.masters[0]!.addText('Formatted master line', {
        name: 'formatted_master_line',
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '445566' },
          transparency: 100,
          width: 0,
          dash: 'sysDot',
        },
      });
      formattedLayout.addPlaceholder('Formatted prompt', {
        name: 'formatted_title_line',
        type: 'title',
        index: 195,
        line: { kind: 'line', color: { kind: 'scheme', value: 'accent6' } },
      });
      const formattedSlide = formatted.addSlide({ masterName: formattedLayout.name });
      formattedSlide.addText('Formatted populated line', {
        placeholder: 'formatted_title_line',
        line: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 25,
          width: 2.25,
          dash: 'dashDot',
        },
      });

      const first = await PptxDocument.open(await formatted.write());
      const formattedSecond = await PptxDocument.open(await first.write());
      expect(formattedSecond.format).toBe(format);
      expect(formattedSecond.formatProfile).toEqual(PRESENTATION_FORMAT_PROFILES[format]);
      expect((formattedSecond.layouts[0]!.shapes.find(
        ({ name }) => name === 'formatted_layout_line',
      ) as ShapeModel).line).toEqual({
        kind: 'line',
        color: { kind: 'scheme', value: 'accent5' },
        width: 1,
        dash: 'solid',
      });
      expect((formattedSecond.masters[0]!.shapes.find(
        ({ name }) => name === 'formatted_master_line',
      ) as ShapeModel).line).toEqual({
        kind: 'line',
        color: { kind: 'srgb', value: '445566' },
        transparency: 100,
        width: 0,
        dash: 'sysDot',
      });
      expect((formattedSecond.slides[0]!.shapes.find(
        ({ name }) => name === 'formatted_title_line',
      ) as ShapeModel).line).toEqual({
        kind: 'line',
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 25,
        width: 2.25,
        dash: 'dashDot',
      });
      expect(validatePackage(formattedSecond.opcPackage).filter(
        ({ severity }) => severity === 'error',
      )).toEqual([]);
    }
  });

  it('creates text arrows across slide layout master and placeholder owners', async () => {
    const document = PptxDocument.create();
    const layout = document.layouts[0]!;
    const master = document.masters[0]!;
    const layoutText = layout.addText('Layout arrows', {
      name: 'layout_arrows',
      arrows: { begin: 'triangle' },
    });
    const masterText = master.addRichText([{ runs: [{ text: 'Master arrows' }] }], {
      name: 'master_arrows',
      arrows: { end: 'arrow' },
    });
    const layoutPlaceholder = layout.addPlaceholder('Layout prompt', {
      name: 'title_arrows',
      type: 'title',
      index: 196,
      x: inches(1),
      y: inches(1),
      width: inches(8),
      height: inches(1),
      arrows: { begin: 'none', end: 'stealth' },
    });

    expect(layout.shapes.find(({ id }) => id === layoutText.id)).toBe(layoutText);
    expect(master.shapes.find(({ id }) => id === masterText.id)).toBe(masterText);
    expect(layout.placeholders.find(({ id }) => id === layoutPlaceholder.id))
      .toBe(layoutPlaceholder);
    expect(layoutText.arrows).toEqual({ begin: 'triangle' });
    expect(masterText.arrows).toEqual({ end: 'arrow' });
    expect(layoutPlaceholder.arrows).toEqual({ begin: 'none', end: 'stealth' });

    const slide = document.addSlide({ masterName: layout.name });
    const materialized = slide.placeholders.find(({ name }) => name === 'title_arrows')!;
    const materializedState = {
      id: materialized.id,
      name: materialized.name,
      transform: materialized.transform,
      placeholder: materialized.placeholder,
    };
    const populated = slide.addText('Populated arrows', {
      placeholder: 'title_arrows',
      line: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent3' },
        width: 2.5,
        dash: 'dashDot',
      },
      arrows: { begin: 'diamond', end: 'oval' },
    });
    expect(populated).toBe(slide.shapes.find(({ id }) => id === materializedState.id));
    expect({
      id: populated.id,
      name: populated.name,
      transform: populated.transform,
      placeholder: populated.placeholder,
    }).toEqual(materializedState);
    expect(populated.arrows).toEqual({ begin: 'diamond', end: 'oval' });
    expect(populated.line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent3' },
      width: 2.5,
      dash: 'dashDot',
    });
    expect(layoutPlaceholder.arrows).toEqual({ begin: 'none', end: 'stealth' });
    expect(materialized).not.toBe(populated);

    const declarative = await document.defineSlideMaster({
      title: 'TEXT-ARROWS',
      objects: [
        {
          kind: 'text',
          text: 'Declarative text arrows',
          options: {
            name: 'declarative_text_arrows',
            arrows: { begin: 'stealth', end: 'triangle' },
          },
        },
        {
          kind: 'placeholder',
          text: 'Declarative prompt',
          options: {
            name: 'declarative_title_arrows',
            type: 'title',
            index: 197,
            arrows: { begin: 'arrow', end: 'none' },
          },
        },
      ],
    });
    const declarativeText = declarative.shapes.find(
      ({ name }) => name === 'declarative_text_arrows',
    ) as ShapeModel;
    const declarativePlaceholder = declarative.placeholders.find(
      ({ name }) => name === 'declarative_title_arrows',
    )!;
    expect(declarativeText.arrows).toEqual({ begin: 'stealth', end: 'triangle' });
    expect(declarativePlaceholder.arrows).toEqual({ begin: 'arrow', end: 'none' });
    const declarativeSlide = document.addSlide({ masterName: declarative.name });
    const declarativePopulated = declarativeSlide.addText('Declarative populated', {
      placeholder: 'declarative_title_arrows',
      arrows: { end: 'diamond' },
    });
    expect(declarativePopulated.arrows).toEqual({ end: 'diamond' });
    expect(declarativePlaceholder.arrows).toEqual({ begin: 'arrow', end: 'none' });

    const duplicate = document.duplicateSlide(document.slides.indexOf(slide));
    const duplicatePopulated = duplicate.shapes.find(
      ({ name }) => name === populated.name,
    ) as ShapeModel;
    duplicatePopulated.arrows = { begin: 'triangle' };
    expect(duplicatePopulated.arrows).toEqual({ begin: 'triangle' });
    expect(populated.arrows).toEqual({ begin: 'diamond', end: 'oval' });

    const beforeRollbackShapes = [...slide.shapes];
    let rolledBack: ShapeModel | undefined;
    expect(() => document.transaction(() => {
      rolledBack = slide.addText('Rolled back arrows', {
        arrows: { begin: 'triangle', end: 'arrow' },
      });
      throw new Error('restore created text arrows');
    })).toThrow('restore created text arrows');
    expect(slide.shapes).toEqual(beforeRollbackShapes);
    expect(slide.shapes.find(({ id }) => id === populated.id)).toBe(populated);
    expect(() => rolledBack!.arrows).toThrow(ModelParseError);

    let signalRead!: () => void;
    let resumeRead!: () => void;
    const readStarted = new Promise<void>((resolve) => { signalRead = resolve; });
    const readPaused = new Promise<void>((resolve) => { resumeRead = resolve; });
    const detachedArrows: { begin: 'triangle' | 'oval'; end: 'arrow' | 'diamond' } = {
      begin: 'triangle',
      end: 'arrow',
    };
    const pendingDetached = document.defineSlideMaster({
      title: 'DETACHED-TEXT-ARROWS',
      objects: [
        {
          kind: 'text',
          text: 'Detached text arrows',
          options: { name: 'detached_text_arrows', arrows: detachedArrows },
        },
        {
          kind: 'image',
          source: {
            async *[Symbol.asyncIterator]() {
              signalRead();
              await readPaused;
              yield sdkPngHeader(1, 1);
            },
          },
        },
      ],
    });
    await readStarted;
    detachedArrows.begin = 'oval';
    detachedArrows.end = 'diamond';
    resumeRead();
    const detachedLayout = await pendingDetached;
    expect((detachedLayout.shapes.find(
      ({ name }) => name === 'detached_text_arrows',
    ) as ShapeModel).arrows).toEqual({ begin: 'triangle', end: 'arrow' });

    const reopened = await PptxDocument.open(await document.write());
    const second = await PptxDocument.open(await reopened.write());
    expect((second.layouts.find(({ name }) => name === layout.name)!.shapes.find(
      ({ name }) => name === 'layout_arrows',
    ) as ShapeModel).arrows).toEqual({ begin: 'triangle' });
    expect((second.masters[0]!.shapes.find(
      ({ name }) => name === 'master_arrows',
    ) as ShapeModel).arrows).toEqual({ end: 'arrow' });
    expect((second.slides.find(({ partUri }) => partUri === slide.partUri)!.shapes.find(
      ({ name }) => name === populated.name,
    ) as ShapeModel).arrows).toEqual({ begin: 'diamond', end: 'oval' });
    expect(validatePackage(second.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      const formattedLayout = formatted.layouts[0]!;
      formattedLayout.addText('Formatted layout arrows', {
        name: 'formatted_layout_arrows',
        arrows: { begin: 'triangle' },
      });
      formatted.masters[0]!.addText('Formatted master arrows', {
        name: 'formatted_master_arrows',
        arrows: { end: 'arrow' },
      });
      formattedLayout.addPlaceholder('Formatted prompt', {
        name: 'formatted_title_arrows',
        type: 'title',
        index: 198,
        arrows: { begin: 'none', end: 'stealth' },
      });
      const formattedSlide = formatted.addSlide({ masterName: formattedLayout.name });
      formattedSlide.addText('Formatted populated arrows', {
        placeholder: 'formatted_title_arrows',
        arrows: { begin: 'diamond', end: 'oval' },
      });

      const first = await PptxDocument.open(await formatted.write());
      const formattedSecond = await PptxDocument.open(await first.write());
      expect(formattedSecond.format).toBe(format);
      expect(formattedSecond.formatProfile).toEqual(PRESENTATION_FORMAT_PROFILES[format]);
      expect((formattedSecond.layouts[0]!.shapes.find(
        ({ name }) => name === 'formatted_layout_arrows',
      ) as ShapeModel).arrows).toEqual({ begin: 'triangle' });
      expect((formattedSecond.masters[0]!.shapes.find(
        ({ name }) => name === 'formatted_master_arrows',
      ) as ShapeModel).arrows).toEqual({ end: 'arrow' });
      expect((formattedSecond.slides[0]!.shapes.find(
        ({ name }) => name === 'formatted_title_arrows',
      ) as ShapeModel).arrows).toEqual({ begin: 'diamond', end: 'oval' });
      expect(validatePackage(formattedSecond.opcPackage).filter(
        ({ severity }) => severity === 'error',
      )).toEqual([]);
    }
  });

  it('rejects invalid declarative text arrows without observable mutation', async () => {
    const document = PptxDocument.create();
    const { output: _beforeOutput, ...before } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    for (const definition of [
      {
        title: 'INVALID-TEXT-ARROWS',
        objects: [{
          kind: 'text',
          text: 'Invalid text arrows',
          options: { arrows: { beginArrowType: 'triangle' } },
        }],
      },
      {
        title: 'INVALID-PLACEHOLDER-ARROWS',
        objects: [{
          kind: 'placeholder',
          options: {
            name: 'invalid_arrows',
            type: 'title',
            arrows: { end: 'bogus' },
          },
        }],
      },
    ]) {
      await expect(document.defineSlideMaster(definition as never)).rejects.toThrow();
      const { output: _afterOutput, ...after } = await sdkPackageSnapshot(document) as {
        readonly output: Uint8Array;
        readonly [key: string]: unknown;
      };
      expect(after).toEqual(before);
    }

    let signalRead!: () => void;
    let resumeRead!: () => void;
    const readStarted = new Promise<void>((resolve) => { signalRead = resolve; });
    const readPaused = new Promise<void>((resolve) => { resumeRead = resolve; });
    const pending = document.defineSlideMaster({
      title: 'INVALID-ASYNC-TEXT-ARROWS',
      objects: [
        {
          kind: 'text',
          text: 'Invalid detached arrows',
          options: { arrows: { begin: 'bogus' } as never },
        },
        {
          kind: 'image',
          source: {
            async *[Symbol.asyncIterator]() {
              signalRead();
              await readPaused;
              yield sdkPngHeader(1, 1);
            },
          },
        },
      ],
    });
    const phase = await Promise.race([
      readStarted.then(() => 'read-started' as const),
      pending.then(() => 'resolved' as const, () => 'rejected' as const),
    ]);
    expect(phase).toBe('read-started');
    const { output: _pausedOutput, ...paused } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(paused).toEqual(before);
    resumeRead();
    await expect(pending).rejects.toThrow(/unsupported/i);
    const { output: _rejectedOutput, ...rejected } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(rejected).toEqual(before);
  });

  it('creates text shadows across slide layout master and placeholder owners', async () => {
    const document = PptxDocument.create();
    const layout = document.layouts[0]!;
    const master = document.masters[0]!;
    const layoutText = layout.addText('Layout shadow', {
      name: 'layout_shadow',
      shadow: {
        kind: 'outer',
        color: { kind: 'scheme', value: 'accent2' },
        rotateWithShape: true,
      },
    });
    const masterText = master.addRichText([{ runs: [{ text: 'Master shadow' }] }], {
      name: 'master_shadow',
      shadow: { kind: 'inner', opacity: 0, blur: 0, angle: 0, distance: 0 },
    });
    const layoutPlaceholder = layout.addPlaceholder('Layout prompt', {
      name: 'title_shadow',
      type: 'title',
      index: 199,
      x: inches(1),
      y: inches(1),
      width: inches(8),
      height: inches(1),
      shadow: { kind: 'outer', color: { kind: 'srgb', value: '112233' } },
    });

    expect(layout.shapes.find(({ id }) => id === layoutText.id)).toBe(layoutText);
    expect(master.shapes.find(({ id }) => id === masterText.id)).toBe(masterText);
    expect(layout.placeholders.find(({ id }) => id === layoutPlaceholder.id))
      .toBe(layoutPlaceholder);
    expect(layoutText.shadow).toMatchObject({
      kind: 'outer',
      color: { kind: 'scheme', value: 'accent2' },
      rotateWithShape: true,
    });
    expect(masterText.shadow).toEqual({
      kind: 'inner',
      color: { kind: 'srgb', value: '000000' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
    });

    const slide = document.addSlide({ masterName: layout.name });
    const materialized = slide.placeholders.find(({ name }) => name === 'title_shadow')!;
    const materializedState = {
      id: materialized.id,
      name: materialized.name,
      transform: materialized.transform,
      placeholder: materialized.placeholder,
    };
    const populated = slide.addText('Populated shadow', {
      placeholder: 'title_shadow',
      shadow: {
        kind: 'inner',
        color: { kind: 'scheme', value: 'accent3' },
        opacity: 0.5,
        blur: 2,
        angle: 45,
        distance: 3,
      },
    });
    expect(populated).toBe(slide.shapes.find(({ id }) => id === materializedState.id));
    expect({
      id: populated.id,
      name: populated.name,
      transform: populated.transform,
      placeholder: populated.placeholder,
    }).toEqual(materializedState);
    expect(populated.shadow).toEqual({
      kind: 'inner',
      color: { kind: 'scheme', value: 'accent3' },
      opacity: 0.5,
      blur: 2,
      angle: 45,
      distance: 3,
    });
    expect(layoutPlaceholder.shadow).toMatchObject({
      kind: 'outer',
      color: { kind: 'srgb', value: '112233' },
    });
    expect(materialized).not.toBe(populated);

    const declarative = await document.defineSlideMaster({
      title: 'TEXT-SHADOWS',
      objects: [
        {
          kind: 'text',
          text: 'Declarative text shadow',
          options: {
            name: 'declarative_text_shadow',
            shadow: { kind: 'outer', color: { kind: 'scheme', value: 'accent4' } },
          },
        },
        {
          kind: 'placeholder',
          text: 'Declarative prompt',
          options: {
            name: 'declarative_title_shadow',
            type: 'title',
            index: 200,
            shadow: { kind: 'inner', opacity: 0.25 },
          },
        },
      ],
    });
    const declarativeText = declarative.shapes.find(
      ({ name }) => name === 'declarative_text_shadow',
    ) as ShapeModel;
    const declarativePlaceholder = declarative.placeholders.find(
      ({ name }) => name === 'declarative_title_shadow',
    )!;
    expect(declarativeText.shadow).toMatchObject({
      kind: 'outer',
      color: { kind: 'scheme', value: 'accent4' },
    });
    expect(declarativePlaceholder.shadow).toMatchObject({ kind: 'inner', opacity: 0.25 });
    const declarativeSlide = document.addSlide({ masterName: declarative.name });
    const declarativePopulated = declarativeSlide.addText('Declarative populated', {
      placeholder: 'declarative_title_shadow',
      shadow: { kind: 'outer', rotateWithShape: true },
    });
    expect(declarativePopulated.shadow).toMatchObject({
      kind: 'outer',
      rotateWithShape: true,
    });
    expect(declarativePlaceholder.shadow).toMatchObject({ kind: 'inner', opacity: 0.25 });

    const duplicate = document.duplicateSlide(document.slides.indexOf(slide));
    const duplicatePopulated = duplicate.shapes.find(
      ({ name }) => name === populated.name,
    ) as ShapeModel;
    duplicatePopulated.shadow = { kind: 'outer', color: { kind: 'scheme', value: 'accent5' } };
    expect(duplicatePopulated.shadow).toMatchObject({
      kind: 'outer',
      color: { kind: 'scheme', value: 'accent5' },
    });
    expect(populated.shadow).toMatchObject({
      kind: 'inner',
      color: { kind: 'scheme', value: 'accent3' },
    });

    const beforeRollbackShapes = [...slide.shapes];
    let rolledBack: ShapeModel | undefined;
    expect(() => document.transaction(() => {
      rolledBack = slide.addText('Rolled back shadow', { shadow: { kind: 'outer' } });
      throw new Error('restore created text shadow');
    })).toThrow('restore created text shadow');
    expect(slide.shapes).toEqual(beforeRollbackShapes);
    expect(slide.shapes.find(({ id }) => id === populated.id)).toBe(populated);
    expect(() => rolledBack!.shadow).toThrow(ModelParseError);

    let signalRead!: () => void;
    let resumeRead!: () => void;
    const readStarted = new Promise<void>((resolve) => { signalRead = resolve; });
    const readPaused = new Promise<void>((resolve) => { resumeRead = resolve; });
    const detachedColor: { kind: 'scheme'; value: 'accent1' | 'accent6' } = {
      kind: 'scheme',
      value: 'accent1',
    };
    const detachedShadow: {
      kind: 'outer';
      color: typeof detachedColor;
      opacity: number;
    } = { kind: 'outer', color: detachedColor, opacity: 0.4 };
    const pendingDetached = document.defineSlideMaster({
      title: 'DETACHED-TEXT-SHADOW',
      objects: [
        {
          kind: 'text',
          text: 'Detached text shadow',
          options: { name: 'detached_text_shadow', shadow: detachedShadow },
        },
        {
          kind: 'image',
          source: {
            async *[Symbol.asyncIterator]() {
              signalRead();
              await readPaused;
              yield sdkPngHeader(1, 1);
            },
          },
        },
      ],
    });
    await readStarted;
    detachedColor.value = 'accent6';
    detachedShadow.opacity = 0.9;
    resumeRead();
    const detachedLayout = await pendingDetached;
    expect((detachedLayout.shapes.find(
      ({ name }) => name === 'detached_text_shadow',
    ) as ShapeModel).shadow).toMatchObject({
      color: { kind: 'scheme', value: 'accent1' },
      opacity: 0.4,
    });

    const reopened = await PptxDocument.open(await document.write());
    const second = await PptxDocument.open(await reopened.write());
    expect((second.layouts.find(({ name }) => name === layout.name)!.shapes.find(
      ({ name }) => name === 'layout_shadow',
    ) as ShapeModel).shadow).toMatchObject({ kind: 'outer', rotateWithShape: true });
    expect((second.masters[0]!.shapes.find(
      ({ name }) => name === 'master_shadow',
    ) as ShapeModel).shadow).toMatchObject({ kind: 'inner', opacity: 0 });
    expect((second.slides.find(({ partUri }) => partUri === slide.partUri)!.shapes.find(
      ({ name }) => name === populated.name,
    ) as ShapeModel).shadow).toMatchObject({ kind: 'inner', opacity: 0.5 });
    expect(validatePackage(second.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      const formattedLayout = formatted.layouts[0]!;
      formattedLayout.addText('Formatted layout shadow', {
        name: 'formatted_layout_shadow',
        shadow: { kind: 'outer', color: { kind: 'scheme', value: 'accent2' } },
      });
      formatted.masters[0]!.addRichText([{ runs: [{ text: 'Formatted master shadow' }] }], {
        name: 'formatted_master_shadow',
        shadow: { kind: 'inner', opacity: 0 },
      });
      formattedLayout.addPlaceholder('Formatted prompt', {
        name: 'formatted_title_shadow',
        type: 'title',
        index: 201,
        shadow: { kind: 'outer' },
      });
      const formattedSlide = formatted.addSlide({ masterName: formattedLayout.name });
      formattedSlide.addText('Formatted populated shadow', {
        placeholder: 'formatted_title_shadow',
        shadow: { kind: 'inner', blur: 0, angle: 0, distance: 0 },
      });

      const first = await PptxDocument.open(await formatted.write());
      const formattedSecond = await PptxDocument.open(await first.write());
      expect(formattedSecond.format).toBe(format);
      expect(formattedSecond.formatProfile).toEqual(PRESENTATION_FORMAT_PROFILES[format]);
      expect((formattedSecond.layouts[0]!.shapes.find(
        ({ name }) => name === 'formatted_layout_shadow',
      ) as ShapeModel).shadow).toMatchObject({ kind: 'outer' });
      expect((formattedSecond.masters[0]!.shapes.find(
        ({ name }) => name === 'formatted_master_shadow',
      ) as ShapeModel).shadow).toMatchObject({ kind: 'inner', opacity: 0 });
      expect((formattedSecond.slides[0]!.shapes.find(
        ({ name }) => name === 'formatted_title_shadow',
      ) as ShapeModel).shadow).toMatchObject({ kind: 'inner', blur: 0, angle: 0, distance: 0 });
      expect(validatePackage(formattedSecond.opcPackage).filter(
        ({ severity }) => severity === 'error',
      )).toEqual([]);
    }
  });

  it('rejects invalid declarative text shadows without observable mutation', async () => {
    const document = PptxDocument.create();
    const { output: _beforeOutput, ...before } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    for (const definition of [
      {
        title: 'INVALID-TEXT-SHADOW',
        objects: [{
          kind: 'text',
          text: 'Invalid text shadow',
          options: { shadow: { type: 'outer' } },
        }],
      },
      {
        title: 'INVALID-PLACEHOLDER-SHADOW',
        objects: [{
          kind: 'placeholder',
          options: {
            name: 'invalid_shadow',
            type: 'title',
            shadow: { kind: 'inner', rotateWithShape: true },
          },
        }],
      },
    ]) {
      await expect(document.defineSlideMaster(definition as never)).rejects.toThrow();
      const { output: _afterOutput, ...after } = await sdkPackageSnapshot(document) as {
        readonly output: Uint8Array;
        readonly [key: string]: unknown;
      };
      expect(after).toEqual(before);
    }

    let signalRead!: () => void;
    let resumeRead!: () => void;
    const readStarted = new Promise<void>((resolve) => { signalRead = resolve; });
    const readPaused = new Promise<void>((resolve) => { resumeRead = resolve; });
    const pending = document.defineSlideMaster({
      title: 'INVALID-ASYNC-TEXT-SHADOW',
      objects: [
        {
          kind: 'text',
          text: 'Invalid detached shadow',
          options: { shadow: { kind: 'outer', distance: 201 } as never },
        },
        {
          kind: 'image',
          source: {
            async *[Symbol.asyncIterator]() {
              signalRead();
              await readPaused;
              yield sdkPngHeader(1, 1);
            },
          },
        },
      ],
    });
    const phase = await Promise.race([
      readStarted.then(() => 'read-started' as const),
      pending.then(() => 'resolved' as const, () => 'rejected' as const),
    ]);
    expect(phase).toBe('read-started');
    const { output: _pausedOutput, ...paused } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(paused).toEqual(before);
    resumeRead();
    await expect(pending).rejects.toThrow();
    const { output: _rejectedOutput, ...rejected } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(rejected).toEqual(before);
  });

  it('creates text hyperlinks across slide layout master and placeholder owners', async () => {
    const document = PptxDocument.create();
    const target = document.addSlide();
    const layout = document.layouts[0]!;
    const master = document.masters[0]!;
    const layoutInput: { url: string; tooltip?: string } = {
      url: 'https://layout.example?a=1&b=2',
      tooltip: 'Layout & link',
    };
    const layoutText = layout.addText('Layout hyperlink', {
      name: 'layout_hyperlink',
      hyperlink: layoutInput,
    });
    const masterText = master.addRichText([{ runs: [{ text: 'Master hyperlink' }] }], {
      name: 'master_hyperlink',
      hyperlink: { slide: document.slides.indexOf(target) + 1, tooltip: '' },
    });
    const layoutPlaceholder = layout.addPlaceholder('Layout hyperlink prompt', {
      name: 'title_hyperlink',
      type: 'title',
      index: 202,
      x: inches(1),
      y: inches(1),
      width: inches(8),
      height: inches(1),
      hyperlink: { url: 'https://prompt.example' },
    });
    layoutInput.url = 'https://changed.example';
    layoutInput.tooltip = 'Changed';

    expect(layout.shapes.find(({ id }) => id === layoutText.id)).toBe(layoutText);
    expect(master.shapes.find(({ id }) => id === masterText.id)).toBe(masterText);
    expect(layout.placeholders.find(({ id }) => id === layoutPlaceholder.id))
      .toBe(layoutPlaceholder);
    expect(layoutText.hyperlink).toEqual({
      url: 'https://layout.example?a=1&b=2',
      tooltip: 'Layout & link',
    });
    expect(masterText.hyperlink).toEqual({
      slide: document.slides.indexOf(target) + 1,
      tooltip: '',
    });
    expect(layoutPlaceholder.hyperlink).toEqual({ url: 'https://prompt.example' });
    expect(document.opcPackage.relationships(layout.partUri).filter(
      ({ type }) => type.endsWith('/hyperlink'),
    )).toHaveLength(2);
    expect(document.opcPackage.relationships(master.partUri).find(
      ({ type, resolvedTarget }) => type.endsWith('/slide') && resolvedTarget === target.partUri,
    )).toBeDefined();

    const slide = document.addSlide({ masterName: layout.name });
    const materialized = slide.placeholders.find(
      ({ name }) => name === 'title_hyperlink',
    ) as ShapeModel;
    const materializedState = {
      id: materialized.id,
      name: materialized.name,
      transform: materialized.transform,
      placeholder: materialized.placeholder,
    };
    expect(materialized.hyperlink).toBeUndefined();
    const populated = slide.addText('Populated hyperlink', {
      placeholder: 'title_hyperlink',
      hyperlink: { slide: document.slides.indexOf(target) + 1, tooltip: 'Target' },
    });
    expect(populated).toBe(slide.shapes.find(({ id }) => id === materializedState.id));
    expect({
      id: populated.id,
      name: populated.name,
      transform: populated.transform,
      placeholder: populated.placeholder,
    }).toEqual(materializedState);
    expect(populated.hyperlink).toEqual({
      slide: document.slides.indexOf(target) + 1,
      tooltip: 'Target',
    });
    expect(layoutPlaceholder.hyperlink).toEqual({ url: 'https://prompt.example' });
    expect(materialized).not.toBe(populated);

    const declarative = await document.defineSlideMaster({
      title: 'TEXT-HYPERLINKS',
      objects: [
        {
          kind: 'text',
          text: 'Declarative text hyperlink',
          options: {
            name: 'declarative_text_hyperlink',
            hyperlink: { url: 'https://declarative.example', tooltip: '' },
          },
        },
        {
          kind: 'placeholder',
          text: 'Declarative hyperlink prompt',
          options: {
            name: 'declarative_title_hyperlink',
            type: 'title',
            index: 203,
            hyperlink: { url: 'https://declarative-prompt.example' },
          },
        },
      ],
    });
    expect((declarative.shapes.find(
      ({ name }) => name === 'declarative_text_hyperlink',
    ) as ShapeModel).hyperlink).toEqual({
      url: 'https://declarative.example',
      tooltip: '',
    });
    expect(declarative.placeholders.find(
      ({ name }) => name === 'declarative_title_hyperlink',
    )?.hyperlink).toEqual({ url: 'https://declarative-prompt.example' });
    const declarativeSlide = document.addSlide({ masterName: declarative.name });
    const declarativePopulated = declarativeSlide.addRichText([{
      runs: [{ text: 'Declarative populated hyperlink' }],
    }], {
      placeholder: 'declarative_title_hyperlink',
      hyperlink: { url: 'https://populated.example' },
    });
    expect(declarativePopulated.hyperlink).toEqual({ url: 'https://populated.example' });

    const beforeRollback = await sdkPackageSnapshot(document);
    let rolledBack: ShapeModel | undefined;
    expect(() => document.transaction(() => {
      rolledBack = slide.addText('Rolled back hyperlink', {
        hyperlink: { url: 'https://rollback.example' },
      });
      throw new Error('restore created text hyperlink');
    })).toThrow('restore created text hyperlink');
    expect(await sdkPackageSnapshot(document)).toEqual(beforeRollback);
    expect(() => rolledBack!.hyperlink).toThrow(ModelParseError);

    let signalRead!: () => void;
    let resumeRead!: () => void;
    const readStarted = new Promise<void>((resolve) => { signalRead = resolve; });
    const readPaused = new Promise<void>((resolve) => { resumeRead = resolve; });
    const detachedHyperlink: { url: string; tooltip?: string } = {
      url: 'https://detached.example',
      tooltip: 'Detached',
    };
    const pendingDetached = document.defineSlideMaster({
      title: 'DETACHED-TEXT-HYPERLINK',
      objects: [
        {
          kind: 'text',
          text: 'Detached text hyperlink',
          options: { name: 'detached_text_hyperlink', hyperlink: detachedHyperlink },
        },
        {
          kind: 'image',
          source: {
            async *[Symbol.asyncIterator]() {
              signalRead();
              await readPaused;
              yield sdkPngHeader(1, 1);
            },
          },
        },
      ],
    });
    await readStarted;
    detachedHyperlink.url = 'https://changed.example';
    detachedHyperlink.tooltip = 'Changed';
    resumeRead();
    const detachedLayout = await pendingDetached;
    expect((detachedLayout.shapes.find(
      ({ name }) => name === 'detached_text_hyperlink',
    ) as ShapeModel).hyperlink).toEqual({
      url: 'https://detached.example',
      tooltip: 'Detached',
    });

    const reopened = await PptxDocument.open(await document.write());
    const second = await PptxDocument.open(await reopened.write());
    expect((second.layouts.find(({ name }) => name === layout.name)!.shapes.find(
      ({ name }) => name === 'layout_hyperlink',
    ) as ShapeModel).hyperlink).toEqual({
      url: 'https://layout.example?a=1&b=2',
      tooltip: 'Layout & link',
    });
    expect((second.masters[0]!.shapes.find(
      ({ name }) => name === 'master_hyperlink',
    ) as ShapeModel).hyperlink).toEqual({
      slide: second.slides.findIndex(({ partUri }) => partUri === target.partUri) + 1,
      tooltip: '',
    });
    expect((second.slides.find(({ partUri }) => partUri === slide.partUri)!.shapes.find(
      ({ name }) => name === populated.name,
    ) as ShapeModel).hyperlink).toEqual({
      slide: second.slides.findIndex(({ partUri }) => partUri === target.partUri) + 1,
      tooltip: 'Target',
    });
    expect(validatePackage(second.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);
  });

  it('preserves text hyperlink targets through duplicate move delete and all formats', async () => {
    const document = PptxDocument.create();
    const source = document.addSlide();
    const target = document.addSlide();
    const layoutLink = document.layouts[0]!.addText('Layout target hyperlink', {
      name: 'layout_target_hyperlink',
      hyperlink: { slide: document.slides.indexOf(target) + 1 },
    });
    const masterLink = document.masters[0]!.addRichText([{
      runs: [{ text: 'Master target hyperlink' }],
    }], {
      name: 'master_target_hyperlink',
      hyperlink: { slide: document.slides.indexOf(target) + 1 },
    });
    const external = source.addText('External hyperlink', {
      hyperlink: { url: 'https://example.com', tooltip: 'Visit' },
    });
    const internal = source.addText('Internal hyperlink', {
      hyperlink: { slide: document.slides.indexOf(target) + 1 },
    });
    const self = source.addRichText([{ runs: [{ text: 'Self hyperlink' }] }], {
      hyperlink: { slide: document.slides.indexOf(source) + 1, tooltip: '' },
    });
    const layoutTargetRelationship = document.opcPackage.relationships(
      document.layouts[0]!.partUri,
    ).find(({ resolvedTarget }) => resolvedTarget === target.partUri)!;
    const masterTargetRelationship = document.opcPackage.relationships(
      document.masters[0]!.partUri,
    ).find(({ resolvedTarget }) => resolvedTarget === target.partUri)!;

    const duplicate = document.duplicateSlide(document.slides.indexOf(source));
    const [duplicateExternal, duplicateInternal, duplicateSelf] = duplicate.shapes as ShapeModel[];
    expect(duplicateExternal!.hyperlink).toEqual(external.hyperlink);
    expect(duplicateInternal!.hyperlink).toEqual({
      slide: document.slides.indexOf(target) + 1,
    });
    expect(duplicateSelf!.hyperlink).toEqual({
      slide: document.slides.indexOf(duplicate) + 1,
      tooltip: '',
    });

    document.moveSlide(document.slides.indexOf(target), 0);
    expect(internal.hyperlink).toEqual({ slide: 1 });
    expect(duplicateInternal!.hyperlink).toEqual({ slide: 1 });
    expect(layoutLink.hyperlink).toEqual({ slide: 1 });
    expect(masterLink.hyperlink).toEqual({ slide: 1 });

    document.deleteSlide(document.slides.indexOf(target));
    expect(internal.hyperlink).toBeUndefined();
    expect(duplicateInternal!.hyperlink).toBeUndefined();
    expect(layoutLink.hyperlink).toBeUndefined();
    expect(masterLink.hyperlink).toBeUndefined();
    expect(external.hyperlink).toEqual({ url: 'https://example.com', tooltip: 'Visit' });
    expect(duplicateExternal!.hyperlink).toEqual(external.hyperlink);
    expect(self.hyperlink).toEqual({
      slide: document.slides.indexOf(source) + 1,
      tooltip: '',
    });
    expect(duplicateSelf!.hyperlink).toEqual({
      slide: document.slides.indexOf(duplicate) + 1,
      tooltip: '',
    });
    const layoutXml = new TextDecoder().decode(document.opcPackage.requirePart(
      document.layouts[0]!.partUri,
    ).bytes);
    const masterXml = new TextDecoder().decode(document.opcPackage.requirePart(
      document.masters[0]!.partUri,
    ).bytes);
    expect(layoutXml).not.toContain(`r:id="${layoutTargetRelationship.id}"`);
    expect(masterXml).not.toContain(`r:id="${masterTargetRelationship.id}"`);
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.layouts[0]!.shapes.find(
      ({ name }) => name === 'layout_target_hyperlink',
    ) as ShapeModel).hyperlink).toBeUndefined();
    expect((reopened.masters[0]!.shapes.find(
      ({ name }) => name === 'master_target_hyperlink',
    ) as ShapeModel).hyperlink).toBeUndefined();

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      const first = formatted.addSlide();
      const secondSlide = formatted.addSlide();
      first.addText('Formatted text hyperlink', {
        hyperlink: { slide: formatted.slides.indexOf(secondSlide) + 1, tooltip: '' },
      });
      formatted.layouts[0]!.addRichText([{ runs: [{ text: 'Formatted layout hyperlink' }] }], {
        name: 'formatted_layout_hyperlink',
        hyperlink: { slide: formatted.slides.indexOf(secondSlide) + 1 },
      });
      const formattedReopened = await PptxDocument.open(await formatted.write());
      expect(formattedReopened.format).toBe(format);
      expect(formattedReopened.formatProfile).toEqual(PRESENTATION_FORMAT_PROFILES[format]);
      expect((formattedReopened.slides[0]!.shapes[0] as ShapeModel).hyperlink)
        .toEqual({ slide: 2, tooltip: '' });
      expect((formattedReopened.layouts[0]!.shapes.find(
        ({ name }) => name === 'formatted_layout_hyperlink',
      ) as ShapeModel).hyperlink).toEqual({ slide: 2 });
      expect(validatePackage(formattedReopened.opcPackage).filter(
        ({ severity }) => severity === 'error',
      )).toEqual([]);
    }
  });

  it('supports rich text run hyperlinks across SDK owners and declarative placeholders', async () => {
    const document = PptxDocument.create();
    const target = document.addSlide();
    const layout = document.layouts[0]!;
    const master = document.masters[0]!;
    const layoutText = layout.addRichText([{
      runs: [{
        text: 'Layout run link',
        style: { hyperlink: { url: 'https://layout-run.example', tooltip: '' } },
      }],
    }], { name: 'layout_run_link' });
    const masterText = master.addRichText([{
      runs: [{
        text: 'Master run link',
        style: { hyperlink: { slide: document.slides.indexOf(target) + 1 } },
      }],
    }], { name: 'master_run_link' });
    const prompt = layout.addPlaceholder([{
      runs: [{
        text: 'Prompt run link',
        style: { hyperlink: { url: 'https://prompt-run.example' } },
      }],
    }], {
      name: 'run_link_prompt',
      type: 'body',
      index: 211,
    });

    const detached = { url: 'https://declarative-run.example', tooltip: 'Detached' };
    let signalRead!: () => void;
    let resumeRead!: () => void;
    const readStarted = new Promise<void>((resolve) => { signalRead = resolve; });
    const readPaused = new Promise<void>((resolve) => { resumeRead = resolve; });
    const pending = document.defineSlideMaster({
      title: 'RUN-HYPERLINKS',
      objects: [
        {
          kind: 'text',
          text: [{
            runs: [{ text: 'Declarative run', style: { hyperlink: detached } }],
          }],
          options: { name: 'declarative_run_link' },
        },
        {
          kind: 'placeholder',
          text: [{
            runs: [{
              text: 'Declarative prompt',
              style: { hyperlink: { url: 'https://declarative-prompt-run.example' } },
            }],
          }],
          options: { name: 'declarative_run_prompt', type: 'body', index: 212 },
        },
        {
          kind: 'image',
          source: {
            async *[Symbol.asyncIterator]() {
              signalRead();
              await readPaused;
              yield sdkPngHeader(1, 1);
            },
          },
        },
      ],
    });
    await readStarted;
    detached.url = 'https://changed.example';
    detached.tooltip = 'Changed';
    resumeRead();
    const declarative = await pending;
    const slide = document.addSlide({ masterName: declarative.name });
    const populated = slide.addRichText([{
      runs: [{
        text: 'Populated run',
        style: { hyperlink: { slide: document.slides.indexOf(target) + 1, tooltip: '' } },
      }],
    }], { placeholder: 'declarative_run_prompt' });

    expect(layoutText.richText[0]!.runs[0]!.style?.hyperlink)
      .toEqual({ url: 'https://layout-run.example', tooltip: '' });
    expect(masterText.richText[0]!.runs[0]!.style?.hyperlink)
      .toEqual({ slide: 1 });
    expect(prompt.richText[0]!.runs[0]!.style?.hyperlink)
      .toEqual({ url: 'https://prompt-run.example' });
    expect((declarative.shapes.find(
      ({ name }) => name === 'declarative_run_link',
    ) as ShapeModel).richText[0]!.runs[0]!.style?.hyperlink).toEqual({
      url: 'https://declarative-run.example',
      tooltip: 'Detached',
    });
    expect(populated.richText[0]!.runs[0]!.style?.hyperlink)
      .toEqual({ slide: document.slides.indexOf(target) + 1, tooltip: '' });
    expect(declarative.placeholders.find(
      ({ name }) => name === 'declarative_run_prompt',
    )?.richText[0]!.runs[0]!.style?.hyperlink)
      .toEqual({ url: 'https://declarative-prompt-run.example' });

    layoutText.richText = [{
      runs: [{
        text: 'Layout run link edited',
        style: { hyperlink: false, bold: true },
      }],
    }];
    expect(layoutText.richText[0]!.runs[0]!.style?.hyperlink).toBeUndefined();

    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.masters[0]!.shapes.find(
      ({ name }) => name === 'master_run_link',
    ) as ShapeModel).richText[0]!.runs[0]!.style?.hyperlink).toEqual({ slide: 1 });
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);
  });

  it('supports rich text line breaks across SDK owners and declarative placeholders', async () => {
    const document = PptxDocument.create();
    const directSlide = document.addSlide();
    const layout = document.layouts[0]!;
    const master = document.masters[0]!;
    const paragraphTexts = (shape: ShapeModel) => shape.richText.map(({ runs }) =>
      runs.map(({ text }) => text));
    const richInput = (prefix: string) => [{
      align: 'center' as const,
      runs: [
        { text: `${prefix} first`, breakLine: true },
        { text: '', breakLine: true },
        { text: `${prefix} last`, softBreakBefore: true },
      ],
    }];

    const slideText = directSlide.addRichText(richInput('Slide'), {
      name: 'slide_break_line',
    });
    const layoutText = layout.addRichText(richInput('Layout'), {
      name: 'layout_break_line',
    });
    const masterText = master.addRichText(richInput('Master'), {
      name: 'master_break_line',
    });
    const layoutPrompt = layout.addPlaceholder(richInput('Layout prompt'), {
      name: 'layout_break_prompt',
      type: 'body',
      index: 221,
    });
    const masterPrompt = master.addPlaceholder(richInput('Master prompt'), {
      name: 'master_break_prompt',
      type: 'body',
      index: 222,
    });

    const declarativeText = [{
      runs: [
        { text: 'Declarative first', breakLine: true },
        { text: 'Declarative last' },
      ],
    }];
    const declarativePrompt = [{
      runs: [
        { text: 'Declarative prompt first', breakLine: true },
        { text: 'Declarative prompt last' },
      ],
    }];
    let signalRead!: () => void;
    let resumeRead!: () => void;
    const readStarted = new Promise<void>((resolve) => { signalRead = resolve; });
    const readPaused = new Promise<void>((resolve) => { resumeRead = resolve; });
    const pending = document.defineSlideMaster({
      title: 'RICH-BREAK-LINES',
      objects: [
        {
          kind: 'text',
          text: declarativeText,
          options: { name: 'declarative_break_line' },
        },
        {
          kind: 'placeholder',
          text: declarativePrompt,
          options: { name: 'declarative_break_prompt', type: 'body', index: 223 },
        },
        {
          kind: 'image',
          source: {
            async *[Symbol.asyncIterator]() {
              signalRead();
              await readPaused;
              yield sdkPngHeader(1, 1);
            },
          },
        },
      ],
    });
    await readStarted;
    declarativeText[0]!.runs[0]!.breakLine = false;
    declarativeText[0]!.runs[0]!.text = 'Changed declarative first';
    declarativePrompt[0]!.runs[0]!.breakLine = false;
    declarativePrompt[0]!.runs[0]!.text = 'Changed declarative prompt first';
    resumeRead();
    const declarative = await pending;
    const populatedSlide = document.addSlide({ masterName: declarative.name });
    const populated = populatedSlide.addRichText([{
      runs: [
        { text: 'Populated first', breakLine: true },
        { text: 'Populated last' },
      ],
    }], { placeholder: 'declarative_break_prompt' });

    expect(paragraphTexts(slideText)).toEqual([
      ['Slide first'],
      [],
      ['Slide last'],
    ]);
    expect(paragraphTexts(layoutText)).toEqual([
      ['Layout first'],
      [],
      ['Layout last'],
    ]);
    expect(paragraphTexts(masterText)).toEqual([
      ['Master first'],
      [],
      ['Master last'],
    ]);
    expect(paragraphTexts(layoutPrompt)).toEqual([
      ['Layout prompt first'],
      [],
      ['Layout prompt last'],
    ]);
    expect(paragraphTexts(masterPrompt)).toEqual([
      ['Master prompt first'],
      [],
      ['Master prompt last'],
    ]);
    expect(paragraphTexts(declarative.shapes.find(
      ({ name }) => name === 'declarative_break_line',
    ) as ShapeModel)).toEqual([
      ['Declarative first'],
      ['Declarative last'],
    ]);
    expect(paragraphTexts(declarative.placeholders.find(
      ({ name }) => name === 'declarative_break_prompt',
    )!)).toEqual([
      ['Declarative prompt first'],
      ['Declarative prompt last'],
    ]);
    expect(paragraphTexts(populated)).toEqual([
      ['Populated first'],
      ['Populated last'],
    ]);
    for (const shape of [
      slideText,
      layoutText,
      masterText,
      layoutPrompt,
      masterPrompt,
      populated,
    ]) {
      expect(shape.richText.flatMap(({ runs }) => runs).some((run) =>
        Object.hasOwn(run, 'breakLine'))).toBe(false);
    }
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    const reopened = await PptxDocument.open(await document.write());
    expect(paragraphTexts(reopened.slides.find(
      ({ partUri }) => partUri === directSlide.partUri,
    )!.shapes.find(({ name }) => name === 'slide_break_line') as ShapeModel)).toEqual([
      ['Slide first'],
      [],
      ['Slide last'],
    ]);
    expect(paragraphTexts(reopened.layouts[0]!.shapes.find(
      ({ name }) => name === 'layout_break_line',
    ) as ShapeModel)).toEqual([
      ['Layout first'],
      [],
      ['Layout last'],
    ]);
    expect(paragraphTexts(reopened.masters[0]!.shapes.find(
      ({ name }) => name === 'master_break_line',
    ) as ShapeModel)).toEqual([
      ['Master first'],
      [],
      ['Master last'],
    ]);
    const reopenedDeclarative = reopened.layouts.find(
      ({ name }) => name === 'RICH-BREAK-LINES',
    )!;
    expect(paragraphTexts(reopenedDeclarative.shapes.find(
      ({ name }) => name === 'declarative_break_line',
    ) as ShapeModel)).toEqual([
      ['Declarative first'],
      ['Declarative last'],
    ]);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);
  });

  it('rejects invalid rich text line breaks across declarative owners without mutation', async () => {
    const document = PptxDocument.create();
    document.addSlide();
    const { output: _beforeOutput, ...before } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    let accessorCalls = 0;
    const accessor = Object.defineProperty({ text: 'Accessor' }, 'breakLine', {
      enumerable: true,
      get() {
        accessorCalls += 1;
        return true;
      },
    });
    for (const [index, text] of [
      [{ runs: [{ text: 'Invalid', breakLine: 'yes' }] }],
      [{ runs: [accessor] }],
    ].entries()) {
      await expect(document.defineSlideMaster({
        title: `INVALID-RICH-BREAK-${index}`,
        objects: [
          { kind: 'text', text: text as never },
          {
            kind: 'placeholder',
            text: text as never,
            options: { name: `invalid_break_prompt_${index}`, type: 'body' },
          },
        ],
      })).rejects.toThrow();
      const { output: _afterOutput, ...after } = await sdkPackageSnapshot(document) as {
        readonly output: Uint8Array;
        readonly [key: string]: unknown;
      };
      expect(after).toEqual(before);
    }
    expect(accessorCalls).toBe(0);
  });

  it('preserves rich text run hyperlinks through duplicate delete and all formats', async () => {
    const document = PptxDocument.create();
    const source = document.addSlide();
    const target = document.addSlide();
    const shape = source.addRichText([{
      runs: [
        { text: 'External', style: { hyperlink: { url: 'https://run.example' } } },
        {
          text: 'Other',
          style: { hyperlink: { slide: document.slides.indexOf(target) + 1 } },
        },
        {
          text: 'Self',
          style: { hyperlink: { slide: document.slides.indexOf(source) + 1, tooltip: '' } },
        },
      ],
    }]);
    const duplicate = document.duplicateSlide(document.slides.indexOf(source));
    const duplicateShape = duplicate.shapes.find(({ id }) => id === shape.id) as ShapeModel;
    expect(duplicateShape.richText[0]!.runs.map((run) => run.style?.hyperlink)).toEqual([
      { url: 'https://run.example' },
      { slide: document.slides.indexOf(target) + 1 },
      { slide: document.slides.indexOf(duplicate) + 1, tooltip: '' },
    ]);

    document.moveSlide(document.slides.indexOf(target), 0);
    expect(shape.richText[0]!.runs[1]!.style?.hyperlink).toEqual({ slide: 1 });
    expect(duplicateShape.richText[0]!.runs[1]!.style?.hyperlink).toEqual({ slide: 1 });
    document.deleteSlide(0);
    expect(shape.richText[0]!.runs.map((run) => run.style?.hyperlink)).toEqual([
      { url: 'https://run.example' },
      undefined,
      { slide: document.slides.indexOf(source) + 1, tooltip: '' },
    ]);
    expect(duplicateShape.richText[0]!.runs[2]!.style?.hyperlink)
      .toEqual({ slide: document.slides.indexOf(duplicate) + 1, tooltip: '' });
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      const first = formatted.addSlide();
      const second = formatted.addSlide();
      first.addRichText([{
        runs: [
          { text: 'URL', style: { hyperlink: { url: 'https://format-run.example' } } },
          { text: 'Slide', style: { hyperlink: { slide: 2, tooltip: '' } } },
        ],
      }]);
      formatted.layouts[0]!.addRichText([{
        runs: [{ text: 'Layout', style: { hyperlink: { slide: 2 } } }],
      }], { name: 'format_layout_run_link' });
      formatted.masters[0]!.addRichText([{
        runs: [{ text: 'Master', style: { hyperlink: { url: 'https://master-run.example' } } }],
      }], { name: 'format_master_run_link' });
      const reopened = await PptxDocument.open(await formatted.write());
      expect(reopened.format).toBe(format);
      expect((reopened.slides[0]!.shapes[0] as ShapeModel).richText[0]!.runs
        .map((run) => run.style?.hyperlink)).toEqual([
        { url: 'https://format-run.example' },
        { slide: reopened.slides.indexOf(
          reopened.slides.find(({ partUri }) => partUri === second.partUri)!,
        ) + 1, tooltip: '' },
      ]);
      expect((reopened.layouts[0]!.shapes.find(
        ({ name }) => name === 'format_layout_run_link',
      ) as ShapeModel).richText[0]!.runs[0]!.style?.hyperlink).toEqual({ slide: 2 });
      expect((reopened.masters[0]!.shapes.find(
        ({ name }) => name === 'format_master_run_link',
      ) as ShapeModel).richText[0]!.runs[0]!.style?.hyperlink)
        .toEqual({ url: 'https://master-run.example' });
      expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);
      expect(first.partUri).toBe(reopened.slides[0]!.partUri);
    }
  });

  it('rejects invalid declarative text hyperlinks without observable mutation', async () => {
    const document = PptxDocument.create();
    document.addSlide();
    const { output: _beforeOutput, ...before } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    for (const definition of [
      {
        title: 'INVALID-TEXT-HYPERLINK',
        objects: [{
          kind: 'text',
          text: 'Invalid text hyperlink',
          options: { hyperlink: { url: 'https://example.com', slide: 1 } },
        }],
      },
      {
        title: 'INVALID-PLACEHOLDER-HYPERLINK',
        objects: [{
          kind: 'placeholder',
          options: {
            name: 'invalid_hyperlink',
            type: 'title',
            hyperlink: { slide: 0 },
          },
        }],
      },
    ]) {
      await expect(document.defineSlideMaster(definition as never)).rejects.toThrow();
      const { output: _afterOutput, ...after } = await sdkPackageSnapshot(document) as {
        readonly output: Uint8Array;
        readonly [key: string]: unknown;
      };
      expect(after).toEqual(before);
    }

    let signalRead!: () => void;
    let resumeRead!: () => void;
    const readStarted = new Promise<void>((resolve) => { signalRead = resolve; });
    const readPaused = new Promise<void>((resolve) => { resumeRead = resolve; });
    const pending = document.defineSlideMaster({
      title: 'INVALID-ASYNC-TEXT-HYPERLINK',
      objects: [
        {
          kind: 'text',
          text: 'Invalid detached hyperlink',
          options: { hyperlink: { slide: 99 } },
        },
        {
          kind: 'image',
          source: {
            async *[Symbol.asyncIterator]() {
              signalRead();
              await readPaused;
              yield sdkPngHeader(1, 1);
            },
          },
        },
      ],
    });
    const phase = await Promise.race([
      readStarted.then(() => 'read-started' as const),
      pending.then(() => 'resolved' as const, () => 'rejected' as const),
    ]);
    expect(phase).toBe('read-started');
    const { output: _pausedOutput, ...paused } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(paused).toEqual(before);
    resumeRead();
    await expect(pending).rejects.toThrow(/out of range/i);
    const { output: _rejectedOutput, ...rejected } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(rejected).toEqual(before);
  });

  it('rejects invalid declarative text lines without observable mutation', async () => {
    const document = PptxDocument.create();
    const { output: _beforeOutput, ...before } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    for (const definition of [
      {
        title: 'INVALID-TEXT-LINE',
        objects: [{
          kind: 'text',
          text: 'Invalid text line',
          options: { line: { color: 'FF0000' } },
        }],
      },
      {
        title: 'INVALID-PLACEHOLDER-LINE',
        objects: [{
          kind: 'placeholder',
          options: {
            name: 'invalid_line',
            type: 'title',
            line: {
              kind: 'line',
              color: { kind: 'srgb', value: 'FF0000' },
              width: 1_585,
            },
          },
        }],
      },
    ]) {
      await expect(document.defineSlideMaster(definition as never)).rejects.toThrow();
      const { output: _afterOutput, ...after } = await sdkPackageSnapshot(document) as {
        readonly output: Uint8Array;
        readonly [key: string]: unknown;
      };
      expect(after).toEqual(before);
    }

    let signalRead!: () => void;
    let resumeRead!: () => void;
    const readStarted = new Promise<void>((resolve) => { signalRead = resolve; });
    const readPaused = new Promise<void>((resolve) => { resumeRead = resolve; });
    const pending = document.defineSlideMaster({
      title: 'INVALID-ASYNC-TEXT-LINE',
      objects: [
        {
          kind: 'text',
          text: 'Invalid detached line',
          options: { line: { kind: 'line' } as never },
        },
        {
          kind: 'image',
          source: {
            async *[Symbol.asyncIterator]() {
              signalRead();
              await readPaused;
              yield sdkPngHeader(1, 1);
            },
          },
        },
      ],
    });
    const phase = await Promise.race([
      readStarted.then(() => 'read-started' as const),
      pending.then(() => 'resolved' as const, () => 'rejected' as const),
    ]);
    expect(phase).toBe('read-started');
    const { output: _pausedOutput, ...paused } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(paused).toEqual(before);
    resumeRead();
    await expect(pending).rejects.toThrow(/color|required/i);
    const { output: _rejectedOutput, ...rejected } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(rejected).toEqual(before);
  });

  it('rejects invalid declarative text fills without observable mutation', async () => {
    const document = PptxDocument.create();
    const { output: _beforeOutput, ...before } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    for (const definition of [
      {
        title: 'INVALID-TEXT-FILL',
        objects: [{
          kind: 'text',
          text: 'Invalid text fill',
          options: { fill: { color: 'FF0000' } },
        }],
      },
      {
        title: 'INVALID-PLACEHOLDER-FILL',
        objects: [{
          kind: 'placeholder',
          options: {
            name: 'invalid_fill',
            type: 'title',
            fill: {
              kind: 'solid',
              color: { kind: 'srgb', value: 'FF0000' },
              transparency: 101,
            },
          },
        }],
      },
    ]) {
      await expect(document.defineSlideMaster(definition as never)).rejects.toThrow();
      const { output: _afterOutput, ...after } = await sdkPackageSnapshot(document) as {
        readonly output: Uint8Array;
        readonly [key: string]: unknown;
      };
      expect(after).toEqual(before);
    }

    let signalRead!: () => void;
    let resumeRead!: () => void;
    const readStarted = new Promise<void>((resolve) => { signalRead = resolve; });
    const readPaused = new Promise<void>((resolve) => { resumeRead = resolve; });
    const pending = document.defineSlideMaster({
      title: 'INVALID-ASYNC-TEXT-FILL',
      objects: [
        {
          kind: 'text',
          text: 'Invalid detached fill',
          options: { fill: { kind: 'solid' } as never },
        },
        {
          kind: 'image',
          source: {
            async *[Symbol.asyncIterator]() {
              signalRead();
              await readPaused;
              yield sdkPngHeader(1, 1);
            },
          },
        },
      ],
    });
    const phase = await Promise.race([
      readStarted.then(() => 'read-started' as const),
      pending.then(() => 'resolved' as const, () => 'rejected' as const),
    ]);
    expect(phase).toBe('read-started');
    const { output: _pausedOutput, ...paused } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(paused).toEqual(before);
    resumeRead();
    await expect(pending).rejects.toThrow(/color|required/i);
    const { output: _rejectedOutput, ...rejected } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(rejected).toEqual(before);
  });

  it('creates text shape preset geometry across public owners and lifecycle', async () => {
    const document = PptxDocument.create();
    const layout = document.layouts[0]!;
    const master = document.masters[0]!;
    const layoutText = layout.addText('Layout geometry', {
      name: 'layout_text_geometry',
      shape: 'ellipse',
    });
    const masterText = master.addRichText([{
      runs: [{ text: 'Master geometry' }],
    }], {
      name: 'master_text_geometry',
      shape: 'star5',
    });
    const layoutPlaceholder = layout.addPlaceholder('Geometry prompt', {
      name: 'geometry_title',
      type: 'title',
      index: 221,
      shape: 'roundRect',
      x: inches(1),
      y: inches(1),
      width: inches(8),
      height: inches(1),
    });
    expect(layoutText.presetType).toBe('ellipse');
    expect(masterText.presetType).toBe('star5');
    expect(layoutPlaceholder.presetType).toBe('roundRect');
    expect(layout.shapes.find(({ id }) => id === layoutText.id)).toBe(layoutText);
    expect(master.shapes.find(({ id }) => id === masterText.id)).toBe(masterText);

    const slide = document.addSlide({ masterName: layout.name });
    const materialized = slide.placeholders.find(({ name }) => name === 'geometry_title')!;
    const materializedState = {
      id: materialized.id,
      name: materialized.name,
      transform: materialized.transform,
      placeholder: materialized.placeholder,
    };
    const populated = slide.addText('Populated geometry', {
      placeholder: 'geometry_title',
      shape: 'diamond',
    });
    expect(populated).toBe(slide.shapes.find(({ id }) => id === materializedState.id));
    expect({
      id: populated.id,
      name: populated.name,
      transform: populated.transform,
      placeholder: populated.placeholder,
    }).toEqual(materializedState);
    expect(populated.presetType).toBe('diamond');
    expect(layoutPlaceholder.presetType).toBe('roundRect');
    expect(materialized).not.toBe(populated);

    const declarative = await document.defineSlideMaster({
      title: 'TEXT-PRESET-GEOMETRY',
      objects: [
        {
          kind: 'text',
          text: 'Declarative geometry',
          options: { name: 'declarative_text_geometry', shape: 'hexagon' },
        },
        {
          kind: 'placeholder',
          text: [{ runs: [{ text: 'Declarative prompt' }] }],
          options: {
            name: 'declarative_geometry_title',
            type: 'title',
            index: 222,
            shape: 'flowChartDecision',
          },
        },
      ],
    });
    expect((declarative.shapes.find(
      ({ name }) => name === 'declarative_text_geometry',
    ) as ShapeModel).presetType).toBe('hexagon');
    expect(declarative.placeholders.find(
      ({ name }) => name === 'declarative_geometry_title',
    )?.presetType).toBe('flowChartDecision');
    const declarativeSlide = document.addSlide({ masterName: declarative.name });
    const declarativePopulated = declarativeSlide.addRichText([{
      runs: [{ text: 'Declarative populated' }],
    }], {
      placeholder: 'declarative_geometry_title',
      shape: 'actionButtonHome',
    });
    expect(declarativePopulated.presetType).toBe('actionButtonHome');

    const duplicate = document.duplicateSlide(document.slides.indexOf(slide));
    const duplicatePopulated = duplicate.shapes.find(
      ({ name }) => name === populated.name,
    ) as ShapeModel;
    expect(duplicatePopulated.presetType).toBe('diamond');
    duplicatePopulated.presetType = 'triangle';
    expect(duplicatePopulated.presetType).toBe('triangle');
    expect(populated.presetType).toBe('diamond');
    document.moveSlide(document.slides.indexOf(duplicate), 0);
    document.moveSlide(0, document.slides.indexOf(slide));

    const beforeRollback = await sdkPackageSnapshot(document);
    let rolledBack: ShapeModel | undefined;
    expect(() => document.transaction(() => {
      rolledBack = slide.addText('Rolled back geometry', { shape: 'cloud' });
      populated.presetType = 'star8';
      throw new Error('restore text preset geometry');
    })).toThrow('restore text preset geometry');
    expect(await sdkPackageSnapshot(document)).toEqual(beforeRollback);
    expect(populated.presetType).toBe('diamond');
    expect(() => rolledBack!.presetType).toThrow(ModelParseError);

    let signalRead!: () => void;
    let resumeRead!: () => void;
    const readStarted = new Promise<void>((resolve) => { signalRead = resolve; });
    const readPaused = new Promise<void>((resolve) => { resumeRead = resolve; });
    const detachedOptions: { shape: 'ellipse' | 'star8' } = { shape: 'ellipse' };
    const pendingDetached = document.defineSlideMaster({
      title: 'DETACHED-TEXT-PRESET-GEOMETRY',
      objects: [
        {
          kind: 'text',
          text: 'Detached geometry',
          options: detachedOptions,
        },
        {
          kind: 'image',
          source: {
            async *[Symbol.asyncIterator]() {
              signalRead();
              await readPaused;
              yield sdkPngHeader(1, 1);
            },
          },
        },
      ],
    });
    await readStarted;
    detachedOptions.shape = 'star8';
    resumeRead();
    const detachedLayout = await pendingDetached;
    expect((detachedLayout.shapes.find(
      ({ name }) => name === 'Text 2',
    ) as ShapeModel).presetType).toBe('ellipse');

    const reopened = await PptxDocument.open(await document.write());
    const second = await PptxDocument.open(await reopened.write());
    expect((second.layouts.find(({ name }) => name === layout.name)!.shapes.find(
      ({ name }) => name === 'layout_text_geometry',
    ) as ShapeModel).presetType).toBe('ellipse');
    expect((second.masters[0]!.shapes.find(
      ({ name }) => name === 'master_text_geometry',
    ) as ShapeModel).presetType).toBe('star5');
    expect((second.slides.find(({ partUri }) => partUri === slide.partUri)!.shapes.find(
      ({ name }) => name === populated.name,
    ) as ShapeModel).presetType).toBe('diamond');
    expect(validatePackage(second.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      formatted.layouts[0]!.addText('Formatted layout geometry', {
        name: 'formatted_layout_geometry',
        shape: 'ellipse',
      });
      formatted.masters[0]!.addText('Formatted master geometry', {
        name: 'formatted_master_geometry',
        shape: 'foldedCorner',
      });
      const formattedSlide = formatted.addSlide();
      formattedSlide.addRichText([{ runs: [{ text: 'Formatted slide geometry' }] }], {
        name: 'formatted_slide_geometry',
        shape: 'star5',
      });
      const formattedReopened = await PptxDocument.open(await formatted.write());
      expect(formattedReopened.format).toBe(format);
      expect((formattedReopened.layouts[0]!.shapes.find(
        ({ name }) => name === 'formatted_layout_geometry',
      ) as ShapeModel).presetType).toBe('ellipse');
      expect((formattedReopened.masters[0]!.shapes.find(
        ({ name }) => name === 'formatted_master_geometry',
      ) as ShapeModel).presetType).toBe('foldedCorner');
      expect((formattedReopened.slides[0]!.shapes.find(
        ({ name }) => name === 'formatted_slide_geometry',
      ) as ShapeModel).presetType).toBe('star5');
      expect(validatePackage(formattedReopened.opcPackage).filter(
        ({ severity }) => severity === 'error',
      )).toEqual([]);
    }
  });

  it('rejects invalid declarative text shape preset geometry without observable mutation', async () => {
    const document = PptxDocument.create();
    const { output: _beforeOutput, ...before } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    for (const definition of [
      {
        title: 'INVALID-TEXT-PRESET-GEOMETRY',
        objects: [{
          kind: 'text',
          text: 'Invalid text geometry',
          options: { shape: 'folderCorner' },
        }],
      },
      {
        title: 'INVALID-PLACEHOLDER-PRESET-GEOMETRY',
        objects: [{
          kind: 'placeholder',
          options: {
            name: 'invalid_geometry',
            type: 'title',
            shape: 'custGeom',
          },
        }],
      },
    ]) {
      await expect(document.defineSlideMaster(definition as never)).rejects.toThrow(TypeError);
      const { output: _afterOutput, ...after } = await sdkPackageSnapshot(document) as {
        readonly output: Uint8Array;
        readonly [key: string]: unknown;
      };
      expect(after).toEqual(before);
    }

    let signalRead!: () => void;
    let resumeRead!: () => void;
    const readStarted = new Promise<void>((resolve) => { signalRead = resolve; });
    const readPaused = new Promise<void>((resolve) => { resumeRead = resolve; });
    const pending = document.defineSlideMaster({
      title: 'INVALID-ASYNC-TEXT-PRESET-GEOMETRY',
      objects: [
        {
          kind: 'text',
          text: 'Invalid detached geometry',
          options: { shape: 'unknown' as never },
        },
        {
          kind: 'image',
          source: {
            async *[Symbol.asyncIterator]() {
              signalRead();
              await readPaused;
              yield sdkPngHeader(1, 1);
            },
          },
        },
      ],
    });
    const phase = await Promise.race([
      readStarted.then(() => 'read-started' as const),
      pending.then(() => 'resolved' as const, () => 'rejected' as const),
    ]);
    expect(phase).toBe('read-started');
    const { output: _pausedOutput, ...paused } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(paused).toEqual(before);
    resumeRead();
    await expect(pending).rejects.toThrow(TypeError);
    const { output: _rejectedOutput, ...rejected } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    expect(rejected).toEqual(before);
  });

  it('creates text shape rectangle radius across public owners and lifecycle', async () => {
    const document = PptxDocument.create();
    const layout = document.layouts[0]!;
    const master = document.masters[0]!;
    const layoutText = layout.addText('Layout rounded text', {
      name: 'layout_rounded_text',
      shape: 'roundRect',
      rectRadius: inches(0.5),
      width: inches(4),
      height: inches(2),
    });
    const masterText = master.addRichText([{
      runs: [{ text: 'Master rounded text' }],
    }], {
      name: 'master_rounded_text',
      shape: 'roundRect',
      rectRadius: inches(0.25),
      width: inches(2),
      height: inches(1),
    });
    const layoutPlaceholder = layout.addPlaceholder('Rounded prompt', {
      name: 'rounded_title',
      type: 'title',
      index: 231,
      shape: 'roundRect',
      rectRadius: inches(0.25),
      x: inches(1),
      y: inches(1),
      width: inches(4),
      height: inches(2),
    });
    const masterPlaceholder = master.addPlaceholder('Master rounded prompt', {
      name: 'master_rounded_title',
      type: 'title',
      index: 232,
      shape: 'roundRect',
      rectRadius: inches(0.5),
      width: inches(4),
      height: inches(2),
    });
    expect(layoutText.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    expect(masterText.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    expect(layoutPlaceholder.adjustments).toEqual([{ name: 'adj', value: 12_500 }]);
    expect(masterPlaceholder.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    expect(layout.shapes.find(({ id }) => id === layoutText.id)).toBe(layoutText);
    expect(master.shapes.find(({ id }) => id === masterText.id)).toBe(masterText);

    const slide = document.addSlide({ masterName: layout.name });
    const materialized = slide.placeholders.find(({ name }) => name === 'rounded_title')!;
    const ownerState = {
      id: materialized.id,
      name: materialized.name,
      transform: materialized.transform,
      placeholder: materialized.placeholder,
    };
    const populated = slide.addText('Populated rounded text', {
      placeholder: 'rounded_title',
      shape: 'roundRect',
      rectRadius: inches(0.5),
      width: inches(1),
      height: inches(1),
    });
    expect(populated).toBe(slide.shapes.find(({ id }) => id === ownerState.id));
    expect({
      id: populated.id,
      name: populated.name,
      transform: populated.transform,
      placeholder: populated.placeholder,
    }).toEqual(ownerState);
    expect(populated.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    expect(layoutPlaceholder.adjustments).toEqual([{ name: 'adj', value: 12_500 }]);
    expect(materialized).not.toBe(populated);
    const slideRich = slide.addRichText([{
      runs: [{ text: 'Slide rounded rich text' }],
    }], {
      name: 'slide_rounded_rich_text',
      shape: 'roundRect',
      rectRadius: inches(0.5),
      width: inches(2),
      height: inches(1),
    });
    const slidePlaceholder = slide.addPlaceholder('Slide rounded prompt', {
      name: 'slide_rounded_prompt',
      type: 'body',
      index: 233,
      shape: 'roundRect',
      rectRadius: inches(0),
      width: inches(2),
      height: inches(1),
    });
    expect(slideRich.adjustments).toEqual([{ name: 'adj', value: 50_000 }]);
    expect(slidePlaceholder.adjustments).toEqual([{ name: 'adj', value: 0 }]);

    const declarative = await document.defineSlideMaster({
      title: 'TEXT-RECT-RADIUS',
      objects: [
        {
          kind: 'text',
          text: 'Declarative rounded text',
          options: {
            name: 'declarative_rounded_text',
            shape: 'roundRect',
            rectRadius: inches(0.25),
            width: inches(2),
            height: inches(1),
          },
        },
        {
          kind: 'placeholder',
          text: [{ runs: [{ text: 'Declarative rounded prompt' }] }],
          options: {
            name: 'declarative_rounded_title',
            type: 'title',
            index: 234,
            shape: 'roundRect',
            rectRadius: inches(0.5),
            width: inches(4),
            height: inches(2),
          },
        },
      ],
    });
    expect((declarative.shapes.find(
      ({ name }) => name === 'declarative_rounded_text',
    ) as ShapeModel).adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    expect(declarative.placeholders.find(
      ({ name }) => name === 'declarative_rounded_title',
    )?.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    const declarativeSlide = document.addSlide({ masterName: declarative.name });
    const declarativePopulated = declarativeSlide.addRichText([{
      runs: [{ text: 'Declarative populated radius' }],
    }], {
      placeholder: 'declarative_rounded_title',
      shape: 'roundRect',
      rectRadius: inches(0.25),
      width: inches(1),
      height: inches(1),
    });
    expect(declarativePopulated.adjustments).toEqual([{ name: 'adj', value: 12_500 }]);

    const duplicate = document.duplicateSlide(document.slides.indexOf(slide));
    const duplicatePopulated = duplicate.shapes.find(
      ({ name }) => name === populated.name,
    ) as ShapeModel;
    expect(duplicatePopulated.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    duplicatePopulated.adjustments = [{ name: 'adj', value: 75_000 }];
    expect(populated.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    document.moveSlide(document.slides.indexOf(duplicate), 0);
    expect(duplicatePopulated.adjustments).toEqual([{ name: 'adj', value: 75_000 }]);

    const beforeRollback = await sdkPackageSnapshot(document);
    expect(() => document.transaction(() => {
      populated.adjustments = [{ name: 'adj', value: 10_000 }];
      populated.setTransform({ width: inches(8), height: inches(4) });
      throw new Error('restore text rectangle radius');
    })).toThrow('restore text rectangle radius');
    expect(await sdkPackageSnapshot(document)).toEqual(beforeRollback);
    expect(populated.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);

    let signalRead!: () => void;
    let resumeRead!: () => void;
    const readStarted = new Promise<void>((resolve) => { signalRead = resolve; });
    const readPaused = new Promise<void>((resolve) => { resumeRead = resolve; });
    const detachedOptions = {
      shape: 'roundRect' as const,
      rectRadius: inches(0.25),
      width: inches(2),
      height: inches(1),
    };
    const pendingDetached = document.defineSlideMaster({
      title: 'DETACHED-TEXT-RECT-RADIUS',
      objects: [
        { kind: 'text', text: 'Detached rounded text', options: detachedOptions },
        {
          kind: 'image',
          source: {
            async *[Symbol.asyncIterator]() {
              signalRead();
              await readPaused;
              yield sdkPngHeader(1, 1);
            },
          },
        },
      ],
    });
    await readStarted;
    detachedOptions.rectRadius = inches(0.75);
    detachedOptions.width = inches(8);
    resumeRead();
    const detachedLayout = await pendingDetached;
    expect((detachedLayout.shapes.find(
      ({ name }) => name === 'Text 2',
    ) as ShapeModel).adjustments).toEqual([{ name: 'adj', value: 25_000 }]);

    const reopened = await PptxDocument.open(await document.write());
    const second = await PptxDocument.open(await reopened.write());
    expect((second.layouts.find(({ name }) => name === layout.name)!.shapes.find(
      ({ name }) => name === 'layout_rounded_text',
    ) as ShapeModel).adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    expect((second.masters[0]!.shapes.find(
      ({ name }) => name === 'master_rounded_text',
    ) as ShapeModel).adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    expect((second.slides.find(({ partUri }) => partUri === slide.partUri)!.shapes.find(
      ({ name }) => name === populated.name,
    ) as ShapeModel).adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    expect(validatePackage(second.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      formatted.layouts[0]!.addText('Formatted layout radius', {
        name: 'formatted_layout_radius',
        shape: 'roundRect',
        rectRadius: inches(0.5),
        width: inches(4),
        height: inches(2),
      });
      formatted.masters[0]!.addRichText([{ runs: [{ text: 'Formatted master radius' }] }], {
        name: 'formatted_master_radius',
        shape: 'roundRect',
        rectRadius: inches(0),
      });
      formatted.addSlide().addText('Formatted slide radius', {
        name: 'formatted_slide_radius',
        shape: 'roundRect',
        rectRadius: inches(0.5),
        width: inches(2),
        height: inches(1),
      });
      const formattedReopened = await PptxDocument.open(await formatted.write());
      expect(formattedReopened.format).toBe(format);
      expect((formattedReopened.layouts[0]!.shapes.find(
        ({ name }) => name === 'formatted_layout_radius',
      ) as ShapeModel).adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
      expect((formattedReopened.masters[0]!.shapes.find(
        ({ name }) => name === 'formatted_master_radius',
      ) as ShapeModel).adjustments).toEqual([{ name: 'adj', value: 0 }]);
      expect((formattedReopened.slides[0]!.shapes.find(
        ({ name }) => name === 'formatted_slide_radius',
      ) as ShapeModel).adjustments).toEqual([{ name: 'adj', value: 50_000 }]);
      expect(validatePackage(formattedReopened.opcPackage).filter(
        ({ severity }) => severity === 'error',
      )).toEqual([]);
    }
  });

  it('rejects invalid declarative text shape rectangle radius without observable mutation', async () => {
    const document = PptxDocument.create();
    const { output: _beforeOutput, ...before } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    let accessorCalls = 0;
    const accessor = Object.defineProperty({ shape: 'roundRect' }, 'rectRadius', {
      enumerable: true,
      get() {
        accessorCalls += 1;
        return inches(0.5);
      },
    });
    for (const definition of [
      {
        title: 'INVALID-TEXT-RECT-RADIUS-SHAPE',
        objects: [{
          kind: 'text',
          text: 'Invalid text radius shape',
          options: { shape: 'ellipse', rectRadius: inches(0.5) },
        }],
      },
      {
        title: 'INVALID-PLACEHOLDER-RECT-RADIUS-RANGE',
        objects: [{
          kind: 'placeholder',
          options: {
            name: 'invalid_radius',
            type: 'title',
            shape: 'roundRect',
            rectRadius: inches(1) + 1,
          },
        }],
      },
      {
        title: 'INVALID-TEXT-RECT-RADIUS-ACCESSOR',
        objects: [{ kind: 'text', text: 'Invalid accessor radius', options: accessor }],
      },
    ]) {
      await expect(document.defineSlideMaster(definition as never)).rejects.toThrow();
      const { output: _afterOutput, ...after } = await sdkPackageSnapshot(document) as {
        readonly output: Uint8Array;
        readonly [key: string]: unknown;
      };
      expect(after).toEqual(before);
    }
    expect(accessorCalls).toBe(0);
  });

  it('preserves text box state across public owners and placeholder lifecycle', async () => {
    const document = PptxDocument.create();
    const layout = document.layouts[0]!;
    const master = document.masters[0]!;
    const layoutText = layout.addText('Layout shape text', {
      name: 'layout_shape_text',
      isTextBox: false,
    });
    const masterText = master.addRichText([{
      runs: [{ text: 'Master text box' }],
    }], {
      name: 'master_text_box',
      isTextBox: true,
    });
    const sourceFalse = layout.addPlaceholder('False prompt', {
      name: 'source_false_text_box',
      type: 'title',
      index: 241,
      isTextBox: false,
      x: inches(1),
      y: inches(1),
      width: inches(4),
      height: inches(1),
    });
    const sourceTrue = layout.addPlaceholder('True prompt', {
      name: 'source_true_text_box',
      type: 'body',
      index: 242,
      isTextBox: true,
      x: inches(1),
      y: inches(2),
      width: inches(4),
      height: inches(2),
    });
    const masterPlaceholder = master.addPlaceholder('Master text box prompt', {
      name: 'master_true_text_box',
      type: 'title',
      index: 243,
      isTextBox: true,
    });
    expect(layoutText.isTextBox).toBe(false);
    expect(masterText.isTextBox).toBe(true);
    expect(sourceFalse.isTextBox).toBe(false);
    expect(sourceTrue.isTextBox).toBe(true);
    expect(masterPlaceholder.isTextBox).toBe(true);

    const layoutPart = document.opcPackage.requirePart(layout.partUri);
    const aliasSource = new TextDecoder().decode(layoutPart.bytes)
      .replace(
        `<p:cNvPr id="${sourceFalse.id}" name="${sourceFalse.name}"/><p:cNvSpPr/>`,
        `<p:cNvPr id="${sourceFalse.id}" name="${sourceFalse.name}"/>`
          + '<p:cNvSpPr txBox="off"/>',
      )
      .replace(
        `<p:cNvPr id="${sourceTrue.id}" name="${sourceTrue.name}"/>`
          + '<p:cNvSpPr txBox="1"/>',
        `<p:cNvPr id="${sourceTrue.id}" name="${sourceTrue.name}"/>`
          + '<p:cNvSpPr txBox="on"/>',
      );
    document.opcPackage.setPart(layout.partUri, aliasSource, layoutPart.contentType);
    expect(sourceFalse.isTextBox).toBe(false);
    expect(sourceTrue.isTextBox).toBe(true);
    const layoutSource = document.opcPackage.requirePart(layout.partUri).bytes.slice();
    const slide = document.addSlide({ masterName: layout.name });
    const materializedFalse = slide.placeholders.find(
      ({ name }) => name === sourceFalse.name,
    )!;
    const materializedTrue = slide.placeholders.find(
      ({ name }) => name === sourceTrue.name,
    )!;
    expect((materializedFalse as ShapeModel).isTextBox).toBe(false);
    expect((materializedTrue as ShapeModel).isTextBox).toBe(true);
    const falseOwner = {
      id: materializedFalse.id,
      name: materializedFalse.name,
      transform: materializedFalse.transform,
      placeholder: materializedFalse.placeholder,
    };
    const trueOwner = {
      id: materializedTrue.id,
      name: materializedTrue.name,
      transform: materializedTrue.transform,
      placeholder: materializedTrue.placeholder,
    };

    const populatedFromFalse = slide.addText('Population keeps false source', {
      placeholder: sourceFalse.name,
      isTextBox: true,
    });
    const populatedFromTrue = slide.addRichText([{
      runs: [{ text: 'Population keeps true source' }],
    }], {
      placeholder: sourceTrue.name,
      isTextBox: false,
    });
    expect(populatedFromFalse.isTextBox).toBe(false);
    expect(populatedFromTrue.isTextBox).toBe(true);
    expect({
      id: populatedFromFalse.id,
      name: populatedFromFalse.name,
      transform: populatedFromFalse.transform,
      placeholder: populatedFromFalse.placeholder,
    }).toEqual(falseOwner);
    expect({
      id: populatedFromTrue.id,
      name: populatedFromTrue.name,
      transform: populatedFromTrue.transform,
      placeholder: populatedFromTrue.placeholder,
    }).toEqual(trueOwner);
    expect(materializedFalse).not.toBe(populatedFromFalse);
    expect(materializedTrue).not.toBe(populatedFromTrue);
    expect(sourceFalse.isTextBox).toBe(false);
    expect(sourceTrue.isTextBox).toBe(true);
    expect(document.opcPackage.requirePart(layout.partUri).bytes).toEqual(layoutSource);

    const slideRich = slide.addRichText([{ runs: [{ text: 'Slide text box' }] }], {
      name: 'slide_text_box',
      isTextBox: true,
    });
    const slidePlaceholder = slide.addPlaceholder('Slide shape prompt', {
      name: 'slide_shape_prompt',
      type: 'body',
      index: 244,
      isTextBox: false,
    });
    expect(slideRich.isTextBox).toBe(true);
    expect(slidePlaceholder.isTextBox).toBe(false);

    const declarative = await document.defineSlideMaster({
      title: 'TEXT-BOX-STATE',
      objects: [
        {
          kind: 'text',
          text: 'Declarative text box',
          options: { name: 'declarative_text_box', isTextBox: true },
        },
        {
          kind: 'placeholder',
          text: [{ runs: [{ text: 'Declarative shape prompt' }] }],
          options: {
            name: 'declarative_shape_prompt',
            type: 'title',
            index: 245,
            isTextBox: false,
          },
        },
        {
          kind: 'placeholder',
          text: 'Declarative text box prompt',
          options: {
            name: 'declarative_text_box_prompt',
            type: 'body',
            index: 246,
            isTextBox: true,
          },
        },
      ],
    });
    expect((declarative.shapes.find(
      ({ name }) => name === 'declarative_text_box',
    ) as ShapeModel).isTextBox).toBe(true);
    expect(declarative.placeholders.find(
      ({ name }) => name === 'declarative_shape_prompt',
    )?.isTextBox).toBe(false);
    expect(declarative.placeholders.find(
      ({ name }) => name === 'declarative_text_box_prompt',
    )?.isTextBox).toBe(true);
    const declarativeSlide = document.addSlide({ masterName: declarative.name });
    expect(declarativeSlide.placeholders.map((shape) => (shape as ShapeModel).isTextBox))
      .toEqual([false, true]);

    let signalRead!: () => void;
    let resumeRead!: () => void;
    const readStarted = new Promise<void>((resolve) => { signalRead = resolve; });
    const readPaused = new Promise<void>((resolve) => { resumeRead = resolve; });
    const detachedOptions = { isTextBox: true };
    const pendingDetached = document.defineSlideMaster({
      title: 'DETACHED-TEXT-BOX-STATE',
      objects: [
        { kind: 'text', text: 'Detached text box', options: detachedOptions },
        {
          kind: 'image',
          source: {
            async *[Symbol.asyncIterator]() {
              signalRead();
              await readPaused;
              yield sdkPngHeader(1, 1);
            },
          },
        },
      ],
    });
    await readStarted;
    detachedOptions.isTextBox = false;
    resumeRead();
    const detachedLayout = await pendingDetached;
    expect((detachedLayout.shapes.find(
      ({ name }) => name === 'Text 2',
    ) as ShapeModel).isTextBox).toBe(true);

    const duplicate = document.duplicateSlide(document.slides.indexOf(slide));
    const duplicateTrue = duplicate.shapes.find(
      ({ name }) => name === populatedFromTrue.name,
    ) as ShapeModel;
    expect(duplicateTrue.isTextBox).toBe(true);
    duplicateTrue.isTextBox = false;
    expect(populatedFromTrue.isTextBox).toBe(true);
    document.moveSlide(document.slides.indexOf(duplicate), 0);
    expect(duplicateTrue.isTextBox).toBe(false);

    const beforeRollback = await sdkPackageSnapshot(document);
    expect(() => document.transaction(() => {
      populatedFromFalse.isTextBox = true;
      populatedFromTrue.isTextBox = false;
      throw new Error('restore public text box state');
    })).toThrow('restore public text box state');
    expect(await sdkPackageSnapshot(document)).toEqual(beforeRollback);
    expect(populatedFromFalse.isTextBox).toBe(false);
    expect(populatedFromTrue.isTextBox).toBe(true);

    const reopened = await PptxDocument.open(await document.write());
    const second = await PptxDocument.open(await reopened.write());
    const secondSlide = second.slides.find(({ partUri }) => partUri === slide.partUri)!;
    expect((secondSlide.shapes.find(
      ({ name }) => name === populatedFromFalse.name,
    ) as ShapeModel).isTextBox).toBe(false);
    expect((secondSlide.shapes.find(
      ({ name }) => name === populatedFromTrue.name,
    ) as ShapeModel).isTextBox).toBe(true);
    expect((second.layouts.find(({ name }) => name === layout.name)!.placeholders.find(
      ({ name }) => name === sourceTrue.name,
    ) as ShapeModel).isTextBox).toBe(true);
    expect(validatePackage(second.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      formatted.layouts[0]!.addPlaceholder('Formatted false prompt', {
        name: 'formatted_false_prompt',
        type: 'title',
        index: 247,
        isTextBox: false,
      });
      formatted.layouts[0]!.addPlaceholder('Formatted true prompt', {
        name: 'formatted_true_prompt',
        type: 'body',
        index: 248,
        isTextBox: true,
      });
      const formattedSlide = formatted.addSlide();
      expect(formattedSlide.placeholders.map((shape) => (shape as ShapeModel).isTextBox))
        .toEqual([false, true]);
      const formattedReopened = await PptxDocument.open(await formatted.write());
      expect(formattedReopened.format).toBe(format);
      expect(formattedReopened.layouts[0]!.placeholders.map(({ isTextBox }) => isTextBox))
        .toEqual([false, true]);
      expect(formattedReopened.slides[0]!.placeholders.map(
        (shape) => (shape as ShapeModel).isTextBox,
      ))
        .toEqual([false, true]);
      expect(validatePackage(formattedReopened.opcPackage).filter(
        ({ severity }) => severity === 'error',
      )).toEqual([]);
    }
  });

  it('rejects invalid declarative text box state without observable mutation', async () => {
    const document = PptxDocument.create();
    const { output: _beforeOutput, ...before } = await sdkPackageSnapshot(document) as {
      readonly output: Uint8Array;
      readonly [key: string]: unknown;
    };
    let accessorCalls = 0;
    const accessor = Object.defineProperty({}, 'isTextBox', {
      enumerable: true,
      get() {
        accessorCalls += 1;
        return true;
      },
    });
    for (const definition of [
      {
        title: 'INVALID-TEXT-BOX-STRING',
        objects: [{ kind: 'text', text: 'Invalid', options: { isTextBox: 'true' } }],
      },
      {
        title: 'INVALID-TEXT-BOX-PLACEHOLDER',
        objects: [{
          kind: 'placeholder',
          options: { name: 'invalid_text_box', type: 'title', isTextBox: 1 },
        }],
      },
      {
        title: 'INVALID-TEXT-BOX-ACCESSOR',
        objects: [{ kind: 'text', text: 'Accessor', options: accessor }],
      },
    ]) {
      await expect(document.defineSlideMaster(definition as never)).rejects.toThrow();
      const { output: _afterOutput, ...after } = await sdkPackageSnapshot(document) as {
        readonly output: Uint8Array;
        readonly [key: string]: unknown;
      };
      expect(after).toEqual(before);
    }
    expect(accessorCalls).toBe(0);
  });

  it('surfaces slide-number compatibility warnings and rejects actual id collisions', async () => {
    const compatibilityProfiles = [
      'powerpoint-2010',
      'powerpoint-current',
      'keynote-current',
      'google-slides-import',
      'libreoffice-current',
    ] as const;
    const native = PptxDocument.create({ firstSlideNumber: 4 });
    native.addSlide().slideNumber = {};
    native.layouts[0]!.slideNumber = {};
    native.masters[0]!.slideNumber = {};
    for (const compatibility of compatibilityProfiles) {
      await native.write({ compatibility });
      expect(native.diagnostics.filter(({ code }) => code.startsWith('SLIDE_NUMBER_')))
        .toEqual([]);
    }

    const slide = native.slides[0]!;
    const part = native.opcPackage.requirePart(slide.partUri);
    const noncanonical = new TextDecoder().decode(part.bytes)
      .replace('<a:t>4</a:t>', '<a:t>999</a:t>');
    native.opcPackage.setPart(slide.partUri, noncanonical, part.contentType);
    await native.write({ compatibility: 'libreoffice-current' });
    expect(native.diagnostics.filter(({ code }) => code.startsWith('SLIDE_NUMBER_')))
      .toEqual([expect.objectContaining({
        severity: 'warning',
        code: 'SLIDE_NUMBER_CACHE_NONCANONICAL',
        partUri: slide.partUri,
        compatibility: 'libreoffice-current',
      })]);

    const collisionPart = native.opcPackage.requirePart(slide.partUri);
    const collision = new TextDecoder().decode(collisionPart.bytes).replace(
      '</p:grpSpPr>',
      '</p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Collision"/></p:nvSpPr></p:sp>',
    );
    native.opcPackage.setPart(slide.partUri, collision, collisionPart.contentType);
    await expect(native.write()).rejects.toBeInstanceOf(ValidationError);
    expect(native.diagnostics).toEqual(expect.arrayContaining([
      expect.objectContaining({
        severity: 'error',
        code: 'SLIDE_NUMBER_SHAPE_ID_COLLISION',
        partUri: slide.partUri,
        compatibility: 'powerpoint-current',
      }),
    ]));
    await expect(native.write({ mode: 'permissive' })).resolves.toBeInstanceOf(Uint8Array);
  });

  it('creates a zero-slide presentation and adds blank slides with the default layout', async () => {
    const document = PptxDocument.create();
    expect(document.format).toBe('pptx');
    expect(document.slides).toHaveLength(0);
    expect(document.masters).toHaveLength(1);
    expect(document.layouts).toHaveLength(1);
    expect(document.layouts[0]?.name).toBe('DEFAULT');
    expect(document.themes).toHaveLength(1);
    expect(new TextDecoder().decode(document.opcPackage.requirePart('/ppt/presentation.xml').bytes)).toContain(
      '<p:sldSz cx="9144000" cy="5143500"/>',
    );

    const first = document.addSlide();
    const second = document.addSlide();
    expect([first.slideId, second.slideId]).toEqual([256, 257]);
    expect(first.relationships.find(({ type }) => type.endsWith('/slideLayout'))?.resolvedTarget).toBe(
      document.layouts[0]?.partUri,
    );
    expect(second.relationships.find(({ type }) => type.endsWith('/slideLayout'))?.resolvedTarget).toBe(
      document.layouts[0]?.partUri,
    );
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides).toHaveLength(2);
    expect(reopened.slides.map(({ slideId }) => slideId)).toEqual([256, 257]);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('exports the complete embedded media API with strict public types', async () => {
    const document = PptxDocument.create();
    document.addSlide();
    const kind: MediaKind = 'audio';
    const source: MediaSource = Uint8Array.of(1, 2, 3);
    const methodSource: Parameters<PptxDocument['addAudio']>[1] = source;
    const exportedSource: MediaSource = methodSource;
    const playback: MediaPlaybackSettings = {
      play: 'auto',
      loop: true,
      hideWhenStopped: true,
      volume: 0.25,
    };
    const options: AddMediaOptions = {
      name: 'Typed narration',
      altText: 'Spoken overview',
      contentType: 'audio/mpeg',
      fileName: 'voice.mp3',
      poster: Uint8Array.of(4, 5, 6),
      posterContentType: 'image/png',
      x: -1,
      y: 0,
      width: 1,
      height: 2,
      ...playback,
      transcode: async (bytes, contentType, mediaKind) => ({
        bytes,
        contentType,
        extension: mediaKind === 'audio' ? '.mp3' : '.mp4',
      }),
    };
    const media: MediaModel = await document.addAudio(0, exportedSource, options);

    expect(kind).toBe('audio');
    expect(media).toBeInstanceOf(MediaModel);
    expect(document.media(0)[0]).toBe(media);
    expect(document.slides[0]!.media[0]).toBe(media);
    expect(document.slides[0]!.shapes[0]).toBe(media);
    const partialPlayback: MediaPlaybackSettings = { loop: true };
    media.settings = partialPlayback;
    media.name = 'Typed narration edited';
    media.altText = undefined;
    media.settings = { play: 'click', volume: 1 };
    media.setTransform({ x: inches(1), y: inches(2), width: inches(3), height: inches(4) });
    expect(await media.replaceSource('https://example.com/typed.mp3')).toBe(media);
    expect(media.externalUrl).toBe('https://example.com/typed.mp3');
    await media.replaceSource(Uint8Array.of(9), { contentType: 'audio/mpeg' });
    expect(media.mediaPartUri).toMatch(/\.mp3$/);
    expect(await media.replacePoster(Uint8Array.of(8), {
      contentType: 'image/gif',
      fileName: 'typed.gif',
    })).toBe(media);
    expect(media.posterPartUri).toMatch(/\.gif$/);
    await media.replacePoster();
    expect(media.posterPartUri).toMatch(/\.png$/);
    expect(document.media(0)[0]).toBe(media);
    expect(media.name).toBe('Typed narration edited');
    expect(media.altText).toBeUndefined();
    expect(media.settings).toEqual({
      play: 'click',
      loop: false,
      hideWhenStopped: false,
      volume: 1,
    });
    expect(media).toMatchObject({
      kind: 'audio',
      transform: {
        x: inches(1),
        y: inches(2),
        width: inches(3),
        height: inches(4),
      },
    });
    const removedBySlide = await document.addVideo(0, 'https://example.com/remove-by-slide.mp4');
    document.slides[0]!.deleteMedia(removedBySlide.id);
    expect(document.media(0)).not.toContain(removedBySlide);
    expect(() => removedBySlide.name).toThrow(/not found/);
    const removedByModel = await document.addVideo(0, 'https://example.com/remove-by-model.mp4');
    removedByModel.remove();
    expect(document.media(0)).not.toContain(removedByModel);
    expect(() => removedByModel.setTransform({ x: inches(1) })).toThrow(/not found/);

    if (false) {
      // @ts-expect-error media name must be a string
      const invalidName: AddMediaOptions = { name: 1 };
      // @ts-expect-error media alt text must be a string
      const invalidAlt: AddMediaOptions = { altText: false };
      // @ts-expect-error media play accepts only click or auto
      const invalidPlay: AddMediaOptions = { play: 'hover' };
      // @ts-expect-error media playback flags must be booleans
      const invalidLoop: AddMediaOptions = { loop: 1 };
      // @ts-expect-error media playback flags must be booleans
      const invalidHide: AddMediaOptions = { hideWhenStopped: 'yes' };
      const invalidTranscode: AddMediaOptions = {
        // @ts-expect-error transcoder output bytes must be a Uint8Array
        transcode: async () => ({ bytes: 'bad', contentType: 'audio/mpeg' }),
      };
      void [invalidName, invalidAlt, invalidPlay, invalidLoop, invalidHide, invalidTranscode];
    }
  });

  it('creates every public media source, MIME family, poster family, and external mode', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'pptx-sdk-media-'));
    const audioPath = join(directory, 'path-audio.mp3');
    await writeFile(audioPath, Uint8Array.of(1, 2, 3));
    try {
      const document = PptxDocument.create();
      document.addSlide();
      const created: Array<{
        readonly model: MediaModel;
        readonly media?: readonly [string, string, Uint8Array];
        readonly poster: readonly [string, string, Uint8Array];
        readonly external?: string;
      }> = [];
      const add = (
        model: MediaModel,
        media: readonly [string, string, Uint8Array] | undefined,
        poster: readonly [string, string, Uint8Array],
        external?: string,
      ): void => {
        created.push({ model, ...(media ? { media } : {}), poster, ...(external ? { external } : {}) });
      };

      add(
        await document.addAudio(0, audioPath, {
          name: 'Audio & "path"',
          poster: Uint8Array.of(31, 32),
          posterContentType: 'image/png',
        }),
        ['audio/mpeg', '.mp3', Uint8Array.of(1, 2, 3)],
        ['image/png', '.png', Uint8Array.of(31, 32)],
      );
      add(
        await document.addAudio(0, 'data:audio/mp4;base64,BAUG', {
          name: 'Data <audio>',
          altText: '',
          poster: new File([Uint8Array.of(33, 34)], 'cover.jpeg', { type: 'image/jpeg' }),
          posterContentType: 'image/jpeg',
        }),
        ['audio/mp4', '.m4a', Uint8Array.of(4, 5, 6)],
        ['image/jpeg', '.jpeg', Uint8Array.of(33, 34)],
      );
      add(
        await document.addAudio(0, Uint8Array.of(7, 8), {
          altText: 'Wave > overview',
          contentType: 'audio/wav',
          poster: 'data:image/gif;base64,IyQ=',
        }),
        ['audio/wav', '.wav', Uint8Array.of(7, 8)],
        ['image/gif', '.gif', Uint8Array.of(35, 36)],
      );
      const oggBuffer = Uint8Array.of(9, 10).buffer;
      add(
        await document.addAudio(0, oggBuffer, { contentType: 'audio/ogg' }),
        ['audio/ogg', '.ogg', Uint8Array.of(9, 10)],
        ['image/png', '.png', SDK_DEFAULT_POSTER_BYTES],
      );
      add(
        await document.addAudio(0, Uint8Array.of(11), {
          contentType: 'audio/mpeg',
          transcode: async () => ({
            bytes: Uint8Array.of(12, 13),
            contentType: 'audio/ogg',
            extension: 'ogg',
          }),
        }),
        ['audio/ogg', '.ogg', Uint8Array.of(12, 13)],
        ['image/png', '.png', SDK_DEFAULT_POSTER_BYTES],
      );
      add(
        await document.addVideo(0, new Blob([Uint8Array.of(14, 15)]), {
          contentType: 'video/mp4',
        }),
        ['video/mp4', '.mp4', Uint8Array.of(14, 15)],
        ['image/png', '.png', SDK_DEFAULT_POSTER_BYTES],
      );
      add(
        await document.addVideo(
          0,
          new File([Uint8Array.of(16, 17)], 'named-video.m4v', { type: 'video/mp4' }),
          { contentType: 'video/mp4' },
        ),
        ['video/mp4', '.m4v', Uint8Array.of(16, 17)],
        ['image/png', '.png', SDK_DEFAULT_POSTER_BYTES],
      );
      add(
        await document.addVideo(
          0,
          new ReadableStream<Uint8Array>({
            start(controller) {
              controller.enqueue(Uint8Array.of(18));
              controller.enqueue(Uint8Array.of(19));
              controller.close();
            },
          }),
          { contentType: 'video/quicktime' },
        ),
        ['video/quicktime', '.mov', Uint8Array.of(18, 19)],
        ['image/png', '.png', SDK_DEFAULT_POSTER_BYTES],
      );
      add(
        await document.addVideo(
          0,
          {
            async *[Symbol.asyncIterator]() {
              yield Uint8Array.of(20);
              yield 21;
            },
          },
          { contentType: 'video/webm' },
        ),
        ['video/webm', '.webm', Uint8Array.of(20, 21)],
        ['image/png', '.png', SDK_DEFAULT_POSTER_BYTES],
      );
      add(
        await document.addAudio(0, 'http://example.com/external.mp3'),
        undefined,
        ['image/png', '.png', SDK_DEFAULT_POSTER_BYTES],
        'http://example.com/external.mp3',
      );
      add(
        await document.addVideo(0, 'https://example.com/external.mp4'),
        undefined,
        ['image/png', '.png', SDK_DEFAULT_POSTER_BYTES],
        'https://example.com/external.mp4',
      );

      const immediate = [...document.media(0)].sort((left, right) => left.shapeId - right.shapeId);
      expect(immediate).toEqual(created.map(({ model }) => model));
      const pictureStates = sdkMediaPictureStates(document, 0);
      expect(pictureStates[0]).toMatchObject({
        name: 'Audio & "path"',
        altText: undefined,
      });
      expect(pictureStates[1]).toMatchObject({ name: 'Data <audio>', altText: '' });
      expect(pictureStates[2]).toMatchObject({ altText: 'Wave > overview' });

      for (const [index, expected] of created.entries()) {
        const { model } = expected;
        if (expected.media) {
          expect(model.mediaPartUri?.endsWith(expected.media[1])).toBe(true);
          expect(document.opcPackage.requirePart(model.mediaPartUri!)).toMatchObject({
            contentType: expected.media[0],
            bytes: expected.media[2],
          });
        } else {
          expect(model.mediaPartUri).toBeUndefined();
          expect(model.externalUrl).toBe(expected.external);
        }
        expect(model.posterPartUri?.endsWith(expected.poster[1])).toBe(true);
        expect(document.opcPackage.requirePart(model.posterPartUri!)).toMatchObject({
          contentType: expected.poster[0],
          bytes: expected.poster[2],
        });
        expect(pictureStates[index]?.relationshipTypes).toEqual(expected.external
          ? [
              `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${model.kind}`,
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            ]
          : [
              `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${model.kind}`,
              'http://schemas.microsoft.com/office/2007/relationships/media',
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            ]);
        expect(pictureStates[index]?.kindTargetMode).toBe(expected.external ? 'External' : 'Internal');
        expect(pictureStates[index]?.kindTarget).toBe(expected.external
          ? expected.external
          : `../media/${model.mediaPartUri!.split('/').at(-1)}`);
        expect(pictureStates[index]?.kindResolvedTarget).toBe(expected.external
          ? undefined
          : model.mediaPartUri);
      }
      expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it('detaches paused media work and leaves invalid SDK calls completely unchanged', async () => {
    const document = PptxDocument.create();
    document.addSlide();
    const sourceChunk = Uint8Array.of(1, 2, 3);
    const posterBytes = Uint8Array.of(4, 5, 6);
    const transcodeResult = {
      bytes: Uint8Array.of(7, 8, 9),
      contentType: 'audio/ogg',
      extension: '.ogg',
    };
    let streamPaused!: () => void;
    const reachedStreamPause = new Promise<void>((resolve) => { streamPaused = resolve; });
    let resumeStream!: () => void;
    const streamGate = new Promise<void>((resolve) => { resumeStream = resolve; });
    let transcodePaused!: () => void;
    const reachedTranscodePause = new Promise<void>((resolve) => { transcodePaused = resolve; });
    let resumeTranscode!: () => void;
    const transcodeGate = new Promise<void>((resolve) => { resumeTranscode = resolve; });
    const source: AsyncIterable<Uint8Array> = {
      async *[Symbol.asyncIterator]() {
        yield sourceChunk;
        streamPaused();
        await streamGate;
      },
    };
    const options: AddMediaOptions = {
      name: 'Detached media',
      altText: 'Original description',
      contentType: 'audio/mpeg',
      poster: posterBytes,
      posterContentType: 'image/png',
      transcode: async (bytes, contentType, kind) => {
        expect(bytes).toEqual(Uint8Array.of(1, 2, 3));
        expect(contentType).toBe('audio/mpeg');
        expect(kind).toBe('audio');
        transcodePaused();
        await transcodeGate;
        return transcodeResult;
      },
    };
    const pending = document.addAudio(0, source, options);

    await reachedStreamPause;
    sourceChunk[0] = 99;
    posterBytes[0] = 99;
    (options as { name?: string }).name = 'Changed media';
    (options as { altText?: string }).altText = 'Changed description';
    resumeStream();
    await reachedTranscodePause;
    resumeTranscode();
    const created = await pending;
    transcodeResult.bytes[0] = 99;
    transcodeResult.contentType = 'audio/mpeg';
    transcodeResult.extension = '.mp3';

    expect(document.opcPackage.requirePart(created.mediaPartUri!)).toMatchObject({
      bytes: Uint8Array.of(7, 8, 9),
      contentType: 'audio/ogg',
    });
    expect(document.opcPackage.requirePart(created.posterPartUri!)).toMatchObject({
      bytes: Uint8Array.of(4, 5, 6),
      contentType: 'image/png',
    });
    expect(sdkMediaPictureStates(document, 0)[0]).toMatchObject({
      name: 'Detached media',
      altText: 'Original description',
    });

    const before = await sdkPackageSnapshot(document);
    let consumed = false;
    const unconsumed: AsyncIterable<Uint8Array> = {
      async *[Symbol.asyncIterator]() {
        consumed = true;
        yield Uint8Array.of(1);
      },
    };
    const unsafeOptions = {};
    Object.defineProperty(unsafeOptions, 'name', {
      get() {
        throw new Error('unsafe getter');
      },
    });
    await expect(document.addAudio(99, unconsumed, unsafeOptions)).rejects.toThrow(/out of range/);
    expect(consumed).toBe(false);
    const invalidCalls: Array<() => Promise<unknown>> = [
      () => document.addAudio(0, null as never),
      () => document.addAudio(0, Uint8Array.of(1), { play: 'hover' } as never),
      () => document.addAudio(0, 'data:audio/mpeg;base64,A===', {}),
      () => document.addVideo(0, Uint8Array.of(1), { poster: 'https://example.com/poster.png' }),
      () => document.addVideo(0, 'https://example.com/video.mp4', {
        transcode: async () => ({ bytes: Uint8Array.of(1), contentType: 'video/mp4' }),
      }),
    ];
    for (const call of invalidCalls) {
      await expect(call()).rejects.toThrow();
      expect(await sdkPackageSnapshot(document)).toEqual(before);
    }
  });

  it('surfaces native media timing diagnostics through public writes', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const media = await document.addAudio(0, Uint8Array.of(1), {
      contentType: 'audio/mpeg',
      play: 'auto',
    });
    const part = document.opcPackage.requirePart(slide.partUri);
    const xml = LosslessXmlDocument.parse(part.bytes);
    xml.removeElement(xml.elements('timing')[0]!);
    document.opcPackage.setPart(part.uri, xml.serialize(), part.contentType);

    await document.write({ mode: 'permissive', compatibility: 'powerpoint-2010' });
    expect(document.diagnostics).toEqual(expect.arrayContaining([
      expect.objectContaining({
        code: 'MEDIA_TIMING_MISSING',
        partUri: slide.partUri,
        message: expect.stringContaining(String(media.shapeId)),
      }),
    ]));
  });

  it('surfaces live slide gradient diagnostics without treating simple fills as gradients', async () => {
    const document = PptxDocument.create();
    const gradientSlide = document.addSlide();
    gradientSlide.background = {
      kind: 'path-gradient',
      path: 'circle',
      stops: [
        { offset: 0, color: 'FFFFFF' },
        { offset: 1, color: '000000' },
      ],
    };
    const solidSlide = document.addSlide();
    solidSlide.background = {
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF3399' },
    };

    await document.write({
      mode: 'permissive',
      compatibility: 'google-slides-import',
    });
    expect(document.diagnostics.filter(({ code }) => code.startsWith('GRADIENT_')))
      .toEqual([expect.objectContaining({
        code: 'GRADIENT_PATH_MAY_DEGRADE',
        partUri: gradientSlide.partUri,
      })]);
  });

  it('sets a live image background from a detached raster source', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const png = sdkPngHeader(16, 9);
    const pending: Promise<void> = document.setSlideBackgroundImage(0, png, {
      contentType: 'image/png',
    });
    png.fill(0);
    await expect(pending).resolves.toBeUndefined();

    expect(slide.background).toEqual({
      kind: 'image',
      contentType: 'image/png',
      bytes: sdkPngHeader(16, 9),
    });
    expect(slide.shapes).toEqual([]);
    expect(slide.relationships.filter(({ type }) => type.endsWith('/image'))).toHaveLength(1);

    slide.background = { kind: 'none' };
    expect(slide.background).toEqual({ kind: 'none' });
    slide.background = {
      kind: 'solid',
      color: { kind: 'srgb', value: '112233' },
      transparency: 25,
    };
    expect(slide.background).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '112233' },
      transparency: 25,
    });
    slide.background = {
      kind: 'linear-gradient',
      angle: 45,
      stops: [
        { offset: 0, color: 'FFFFFF' },
        { offset: 1, color: '000000' },
      ],
    };
    expect(slide.background?.kind).toBe('linear-gradient');
    slide.background = undefined;
    expect(slide.background).toBeUndefined();

    if (false) {
      const source: RasterImageSource = sdkPngHeader(1, 1);
      const options: SetSlideBackgroundImageOptions = {
        contentType: 'image/png',
        signal: new AbortController().signal,
      };
      const result: Promise<void> = document.setSlideBackgroundImage(0, source, options);
      void result;
    }
  });

  it('leaves the document unchanged when background source loading or assignment fails', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const snapshot = () => ({
      parts: document.opcPackage.parts.map(({ uri, contentType, bytes }) => ({
        uri,
        contentType,
        bytes: new Uint8Array(bytes),
      })),
      graph: document.opcPackage.graph,
      journal: document.opcPackage.mutations.map((mutation) => ({ ...mutation })),
      background: slide.background,
    });
    const expected = snapshot();
    let sourceReads = 0;
    const countedSource: RasterImageSource = {
      async *[Symbol.asyncIterator]() {
        sourceReads += 1;
        yield sdkPngHeader(1, 1);
      },
    };
    let accessorReads = 0;
    const accessor = Object.defineProperty({}, 'contentType', {
      enumerable: true,
      get() {
        accessorReads += 1;
        return 'image/png';
      },
    });
    const invalidBeforeRead: readonly (() => Promise<void>)[] = [
      () => document.setSlideBackgroundImage(1, countedSource),
      () => document.setSlideBackgroundImage(0, countedSource, null as never),
      () => document.setSlideBackgroundImage(0, countedSource, accessor as never),
      () => document.setSlideBackgroundImage(
        0,
        countedSource,
        Object.create({ contentType: 'image/png' }),
      ),
      () => document.setSlideBackgroundImage(0, countedSource, { unknown: true } as never),
      () => document.setSlideBackgroundImage(0, countedSource, { signal: {} } as never),
    ];
    for (const call of invalidBeforeRead) {
      await expect(call()).rejects.toBeInstanceOf(Error);
      expect(snapshot()).toEqual(expected);
    }
    expect(sourceReads).toBe(0);
    expect(accessorReads).toBe(0);

    const controller = new AbortController();
    const reason = new Error('abort background image');
    controller.abort(reason);
    const invalidAfterRead: readonly (() => Promise<void>)[] = [
      () => document.setSlideBackgroundImage(0, Uint8Array.of(1, 2, 3)),
      () => document.setSlideBackgroundImage(0, sdkPngHeader(1, 1), {
        contentType: 'image/gif',
      }),
      () => document.setSlideBackgroundImage(0, sdkPngHeader(1, 1), {
        signal: controller.signal,
      }),
    ];
    for (const call of invalidAfterRead) {
      await expect(call()).rejects.toBeInstanceOf(Error);
      expect(snapshot()).toEqual(expected);
    }

    const part = document.opcPackage.requirePart(slide.partUri);
    document.opcPackage.setPart(
      slide.partUri,
      new TextDecoder().decode(part.bytes).replace('<p:cSld>', '<p:cSld/><p:cSld>'),
      part.contentType,
    );
    const malformed = snapshot();
    await expect(document.setSlideBackgroundImage(0, sdkPngHeader(1, 1)))
      .rejects.toThrow(ModelParseError);
    expect(snapshot()).toEqual(malformed);
  });

  it('round-trips image, none, solid, and gradient backgrounds twice in all six formats', async () => {
    const png = sdkPngHeader(32, 18);
    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const document = PptxDocument.create({ format });
      document.addSlide();
      await document.setSlideBackgroundImage(0, png);
      document.addSlide().background = { kind: 'none' };
      document.addSlide().background = {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 25,
      };
      document.addSlide().background = {
        kind: 'linear-gradient',
        angle: 45,
        stops: [
          { offset: 0, color: 'FFFFFF' },
          { offset: 1, color: '000000' },
        ],
      };

      const first = await PptxDocument.open(await document.write());
      const second = await PptxDocument.open(await first.write());
      expect(first.formatProfile).toEqual(PRESENTATION_FORMAT_PROFILES[format]);
      expect(second.formatProfile).toEqual(PRESENTATION_FORMAT_PROFILES[format]);
      expect(second.slides[0]!.background).toEqual({
        kind: 'image',
        contentType: 'image/png',
        bytes: png,
      });
      expect(second.slides[0]!.shapes).toEqual([]);
      const imageRelationships = second.slides[0]!.relationships.filter(({ type }) =>
        type.endsWith('/image'));
      expect(imageRelationships).toHaveLength(1);
      expect(second.opcPackage.requirePart(imageRelationships[0]!.resolvedTarget!)).toMatchObject({
        contentType: 'image/png',
        bytes: png,
      });
      const xml = new TextDecoder().decode(
        second.opcPackage.requirePart(second.slides[0]!.partUri).bytes,
      );
      expect(xml).toContain(
        `<a:blipFill><a:blip r:embed="${imageRelationships[0]!.id}"/>`
          + '<a:stretch><a:fillRect/></a:stretch></a:blipFill>',
      );
      expect(second.slides[1]!.background).toEqual({ kind: 'none' });
      expect(second.slides[2]!.background).toEqual({
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 25,
      });
      expect(second.slides[3]!.background?.kind).toBe('linear-gradient');
      expect(validatePackage(second.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);
    }
  });

  it('validates every native slide background across all compatibility profiles', async () => {
    const document = PptxDocument.create();
    document.addSlide();
    await document.setSlideBackgroundImage(0, sdkPngHeader(32, 18));
    document.addSlide().background = { kind: 'none' };
    document.addSlide().background = {
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF3399' },
      transparency: 50,
    };
    document.addSlide().background = {
      kind: 'linear-gradient',
      angle: 45,
      stops: [
        { offset: 0, color: 'FFFFFF' },
        { offset: 1, color: '000000' },
      ],
    };

    for (const compatibility of [
      'powerpoint-2010',
      'powerpoint-current',
      'keynote-current',
      'google-slides-import',
      'libreoffice-current',
    ] as const) {
      await document.write({ mode: 'permissive', compatibility });
      expect(document.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
    }
  });

  it('round-trips canonical audio and video twice in all six presentation formats', async () => {
    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const document = PptxDocument.create({ format });
      document.addSlide();
      const audio = await document.addAudio(0, 'data:audio/mpeg;base64,AQIDBA==', {
        name: 'Audio & narration',
        altText: 'Spoken overview',
        poster: 'data:image/png;base64,BQYH',
        x: -1,
        y: 0,
        width: 1,
        height: 2,
        play: 'auto',
        loop: true,
        hideWhenStopped: true,
        volume: 0.25,
      });
      const video = await document.addVideo(0, Uint8Array.of(8, 9, 10), {
        name: 'Video "overview"',
        altText: '',
        contentType: 'video/mp4',
        fileName: 'overview.m4v',
        poster: 'data:image/jpeg;base64,CwwN',
        x: 3,
        y: 4,
        width: 5,
        height: 6,
      });
      audio.name = 'Audio & narration edited';
      audio.altText = 'Edited spoken overview';
      audio.settings = { play: 'click', volume: 1 };
      audio.settings = { play: 'auto', loop: true, hideWhenStopped: true, volume: 0.25 };
      audio.setTransform({
        x: inches(-2),
        y: inches(1),
        width: inches(2),
        height: inches(3),
      });
      await audio.replaceSource('https://example.com/temporary.mp3');
      await audio.replaceSource('data:audio/mpeg;base64,AQIDBA==');
      await video.replacePoster(Uint8Array.of(14), { contentType: 'image/gif' });
      await video.replacePoster('data:image/jpeg;base64,CwwN');

      const duplicate = document.duplicateSlide(0);
      const duplicateAudio = duplicate.media[0]!;
      const duplicateVideo = duplicate.media[1]!;
      expect(duplicateAudio.mediaPartUri).toBe(audio.mediaPartUri);
      expect(duplicateVideo.mediaPartUri).toBe(video.mediaPartUri);
      await duplicateAudio.replaceSource(Uint8Array.of(14), { contentType: 'audio/ogg' });
      await duplicateAudio.replacePoster(Uint8Array.of(15), { contentType: 'image/gif' });
      duplicateVideo.remove();
      document.moveSlide(1, 0);
      expect(document.slides[0]).toBe(duplicate);
      expect(duplicate.media[0]).toBe(duplicateAudio);
      document.moveSlide(0, 1);
      const reopened = await PptxDocument.open(await document.write());
      reopened.media(0)[0]!.settings = {
        play: 'click',
        loop: false,
        hideWhenStopped: false,
        volume: 0.75,
      };
      reopened.media(0)[1]!.settings = {
        play: 'auto',
        loop: true,
        hideWhenStopped: true,
        volume: 0,
      };
      const second = await PptxDocument.open(await reopened.write());

      expect(reopened.format).toBe(format);
      expect(second.format).toBe(format);
      expect(second.formatProfile).toEqual(PRESENTATION_FORMAT_PROFILES[format]);
      const media = [...second.media(0)].sort((left, right) => left.shapeId - right.shapeId);
      const duplicateMedia = second.media(1);
      expect(media).toMatchObject([
        {
          kind: 'audio',
          mediaPartUri: audio.mediaPartUri,
          posterPartUri: audio.posterPartUri,
          settings: { play: 'click', loop: false, hideWhenStopped: false, volume: 0.75 },
        },
        {
          kind: 'video',
          mediaPartUri: video.mediaPartUri,
          posterPartUri: video.posterPartUri,
          settings: { play: 'auto', loop: true, hideWhenStopped: true, volume: 0 },
        },
      ]);
      expect(second.opcPackage.requirePart(media[0]!.mediaPartUri!)).toMatchObject({
        contentType: 'audio/mpeg',
        bytes: Uint8Array.of(1, 2, 3, 4),
      });
      expect(media[0]!.mediaPartUri).toMatch(/\.mp3$/);
      expect(second.opcPackage.requirePart(media[0]!.posterPartUri!)).toMatchObject({
        contentType: 'image/png',
        bytes: Uint8Array.of(5, 6, 7),
      });
      expect(media[0]!.posterPartUri).toMatch(/\.png$/);
      expect(second.opcPackage.requirePart(media[1]!.mediaPartUri!)).toMatchObject({
        contentType: 'video/mp4',
        bytes: Uint8Array.of(8, 9, 10),
      });
      expect(media[1]!.mediaPartUri).toMatch(/\.m4v$/);
      expect(second.opcPackage.requirePart(media[1]!.posterPartUri!)).toMatchObject({
        contentType: 'image/jpeg',
        bytes: Uint8Array.of(11, 12, 13),
      });
      expect(media[1]!.posterPartUri).toMatch(/\.jpg$/);
      expect(duplicateMedia).toHaveLength(1);
      expect(second.opcPackage.requirePart(duplicateMedia[0]!.mediaPartUri!)).toMatchObject({
        contentType: 'audio/ogg',
        bytes: Uint8Array.of(14),
      });
      expect(second.opcPackage.requirePart(duplicateMedia[0]!.posterPartUri!)).toMatchObject({
        contentType: 'image/gif',
        bytes: Uint8Array.of(15),
      });
      expect(duplicateMedia[0]!.mediaPartUri).not.toBe(media[0]!.mediaPartUri);
      expect(duplicateMedia[0]!.posterPartUri).not.toBe(media[0]!.posterPartUri);
      expect(sdkMediaPictureStates(second, 0)).toEqual([
        {
          shapeId: media[0]!.shapeId,
          kind: 'audio',
          name: 'Audio & narration edited',
          altText: 'Edited spoken overview',
          x: inches(-2),
          y: inches(1),
          width: inches(2),
          height: inches(3),
          hasOfficeMedia: true,
          relationshipTypes: [
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            'http://schemas.microsoft.com/office/2007/relationships/media',
          ],
          kindTargetMode: 'Internal',
          kindTarget: `../media/${media[0]!.mediaPartUri!.split('/').at(-1)}`,
          kindResolvedTarget: media[0]!.mediaPartUri,
        },
        {
          shapeId: media[1]!.shapeId,
          kind: 'video',
          name: 'Video "overview"',
          altText: '',
          x: 3,
          y: 4,
          width: 5,
          height: 6,
          hasOfficeMedia: true,
          relationshipTypes: [
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video',
            'http://schemas.microsoft.com/office/2007/relationships/media',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
          ],
          kindTargetMode: 'Internal',
          kindTarget: `../media/${media[1]!.mediaPartUri!.split('/').at(-1)}`,
          kindResolvedTarget: media[1]!.mediaPartUri,
        },
      ]);
      expect(sdkNativeMediaTimingState(second, 0, media[0]!.shapeId)).toMatchObject({
        kind: 'audio',
        settings: { play: 'click', loop: false, hideWhenStopped: false, volume: 0.75 },
        commands: ['playFrom(0.0)'],
      });
      expect(sdkNativeMediaTimingState(second, 0, media[1]!.shapeId)).toMatchObject({
        kind: 'video',
        settings: { play: 'auto', loop: true, hideWhenStopped: true, volume: 0 },
        commands: ['playFrom(0.0)', 'togglePause'],
      });
      expect(sdkNativeMediaTimingState(second, 1, duplicateMedia[0]!.shapeId)).toMatchObject({
        kind: 'audio',
        settings: { play: 'auto', loop: true, hideWhenStopped: true, volume: 0.25 },
        commands: ['playFrom(0.0)'],
      });
      for (const [slideIndex, expectedTargets] of [
        [0, media.map(({ shapeId }) => shapeId)],
        [1, duplicateMedia.map(({ shapeId }) => shapeId)],
      ] as const) {
        const timing = sdkSlideTimingState(second, slideIndex);
        expect(new Set(timing.ids).size).toBe(timing.ids.length);
        expect([...new Set(timing.targets)].sort((left, right) => left - right))
          .toEqual([...expectedTargets].sort((left, right) => left - right));
      }
      await second.write({ mode: 'permissive', compatibility: 'powerpoint-current' });
      expect(second.diagnostics.filter(({ code }) => code.startsWith('MEDIA_TIMING_'))).toEqual([]);
      expect(validatePackage(second.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);
    }
  });

  it('exports the embedded raster image API with strict public types', async () => {
    const document = PptxDocument.create();
    const contentType: RasterImageContentType = 'image/png';
    const options: AddImageOptions = {
      contentType,
      width: inches(2),
      height: inches(1),
    };
    const slide = document.addSlide();
    const image: ImageModel = slide.addImage(new Uint8Array([1]), options);

    expect(image).toBeInstanceOf(ImageModel);
    expect(image.sourcePartUri).toMatch(/\/ppt\/media\/image\d+\.png$/);
    expect(slide.shapes[0]).toBe(image);
    expect(document.opcPackage.requirePart(image.sourcePartUri!)).toMatchObject({
      contentType: 'image/png',
      bytes: new Uint8Array([1]),
    });
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedImage = reopened.slides[0]!.shapes[0] as ImageModel;
    expect(reopenedImage).toBeInstanceOf(ImageModel);
    expect(reopenedImage.transform).toMatchObject({
      width: inches(2),
      height: inches(1),
    });
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      const formattedSlide = formatted.addSlide();
      formattedSlide.addImage(new Uint8Array([1]), { contentType: 'image/png' });
      formattedSlide.addImage(new Uint8Array([2]), { contentType: 'image/jpeg' });
      formattedSlide.addImage(new Uint8Array([3]), { contentType: 'image/gif' });
      const formattedReopened = await PptxDocument.open(await formatted.write());
      expect(formattedReopened.format).toBe(format);
      expect(validatePackage(formattedReopened.opcPackage)
        .filter(({ severity }) => severity === 'error')).toEqual([]);
    }

    if (false) {
      // @ts-expect-error raster image content types exclude SVG
      const svgType: RasterImageContentType = 'image/svg+xml';
      // @ts-expect-error embedded raster image options require contentType
      const missingType: AddImageOptions = {};
      // @ts-expect-error model image creation excludes path loading
      const pathOptions: AddImageOptions = { contentType, path: './image.png' };
      // @ts-expect-error model image creation excludes data-URI loading
      const dataOptions: AddImageOptions = { contentType, data: 'data:image/png;base64,AQ==' };
      void [svgType, missingType, pathOptions, dataOptions];
    }
  });

  it('exports the low-level embedded SVG image creation types', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const contentType: SvgImageContentType = 'image/svg+xml';
    const options: AddSvgImageOptions = {
      name: 'Typed SVG',
      width: inches(2),
      height: inches(1),
    };
    const image: ImageModel = slide.addSvgImage(
      new Uint8Array([60, 115, 118, 103, 47, 62]),
      sdkPngHeader(1, 1),
      options,
    );

    expect(contentType).toBe('image/svg+xml');
    expect(slide.shapes[0]).toBe(image);
    expect(image.sourcePartUri).toMatch(/\/ppt\/media\/image\d+\.png$/);
    expect(document.opcPackage.parts.some(({ uri, contentType: type }) =>
      uri.endsWith('.svg') && type === contentType)).toBe(true);
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides[0]!.shapes[0]).toBeInstanceOf(ImageModel);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    if (false) {
      // @ts-expect-error low-level SVG options have a fixed content type
      const selectedType: AddSvgImageOptions = { contentType };
      // @ts-expect-error canonical SVG content type excludes LibreOffice's import spelling
      const invalidType: SvgImageContentType = 'image/svg';
      void [selectedType, invalidType];
    }
  });

  it('creates raster and SVG image hyperlinks through the public SDK surface', async () => {
    const document = PptxDocument.create();
    const source = document.addSlide();
    document.addSlide();
    const rasterLink = {
      url: 'https://sdk-images.example?a=1&b=2',
      tooltip: 'SDK raster',
    };
    const rasterPending = document.addImage(0, sdkPngHeader(16, 9), {
      name: 'SDK raster hyperlink',
      hyperlink: rasterLink,
    });
    rasterLink.url = 'https://mutated.example';
    rasterLink.tooltip = 'Mutated';
    const raster = await rasterPending;

    const svgLink: { slide: number; tooltip?: string } = { slide: 2, tooltip: '' };
    const svgPending = document.addImage(0, sdkSvg(16, 9), {
      fallback: sdkPngHeader(1, 1),
      name: 'SDK SVG hyperlink',
      hyperlink: svgLink,
    });
    svgLink.slide = 1;
    delete svgLink.tooltip;
    const svg = await svgPending;

    expect(raster.hyperlink).toEqual({
      url: 'https://sdk-images.example?a=1&b=2',
      tooltip: 'SDK raster',
    });
    expect(svg.hyperlink).toEqual({ slide: 2, tooltip: '' });
    expect(svg.isSvg).toBe(true);
    expect(source.shapes).toEqual([raster, svg]);
    const sourceXml = new TextDecoder().decode(
      document.opcPackage.requirePart(source.partUri).bytes,
    );
    expect(sourceXml).toContain('tooltip="SDK raster"');
    expect(sourceXml).toContain('tooltip="" action="ppaction://hlinksldjump"');

    const beforeInvalid = {
      parts: document.opcPackage.parts.map(({ uri, bytes }) => ({
        uri,
        bytes: new Uint8Array(bytes),
      })),
      graph: document.opcPackage.graph,
      mutations: [...document.opcPackage.mutations],
      shapes: source.shapes,
    };
    await expect(document.addImage(0, sdkPngHeader(1, 1), {
      hyperlink: { slide: 99 },
    })).rejects.toThrow(/out of range/);
    expect({
      parts: document.opcPackage.parts.map(({ uri, bytes }) => ({
        uri,
        bytes: new Uint8Array(bytes),
      })),
      graph: document.opcPackage.graph,
      mutations: [...document.opcPackage.mutations],
      shapes: source.shapes,
    }).toEqual(beforeInvalid);

    const reopened = await PptxDocument.open(await document.write());
    const [reopenedRaster, reopenedSvg] = reopened.slides[0]!.shapes as ImageModel[];
    expect(reopenedRaster!.hyperlink).toEqual({
      url: 'https://sdk-images.example?a=1&b=2',
      tooltip: 'SDK raster',
    });
    expect(reopenedSvg!.hyperlink).toEqual({ slide: 2, tooltip: '' });
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);
  });

  it('binds async image hyperlink targets by slide identity before reading the source', async () => {
    const deferredSource = () => {
      let release!: () => void;
      const gate = new Promise<void>((resolve) => {
        release = resolve;
      });
      const source: RasterImageSource = {
        async *[Symbol.asyncIterator]() {
          await gate;
          yield sdkPngHeader(1, 1);
        },
      };
      return { source, release };
    };

    const reordered = PptxDocument.create();
    const owner = reordered.addSlide();
    const target = reordered.addSlide();
    reordered.addSlide();
    const reorderGate = deferredSource();
    const pendingReorder = reordered.addImage(0, reorderGate.source, {
      hyperlink: { slide: 2 },
    });
    reordered.moveSlide(1, 2);
    reorderGate.release();
    const reorderedImage = await pendingReorder;
    expect(reorderedImage.hyperlink).toEqual({ slide: 3 });
    expect(owner.relationships.some(({ resolvedTarget }) =>
      resolvedTarget === target.partUri)).toBe(true);

    const deleted = PptxDocument.create();
    deleted.addSlide();
    deleted.addSlide();
    deleted.addSlide();
    const deleteGate = deferredSource();
    const pendingDelete = deleted.addImage(0, deleteGate.source, {
      hyperlink: { slide: 2 },
    });
    deleted.deleteSlide(1);
    deleted.addSlide();
    const beforeRejectedCommit = await sdkPackageSnapshot(deleted);
    deleteGate.release();
    await expect(pendingDelete).rejects.toThrow(/target was deleted/);
    expect(await sdkPackageSnapshot(deleted)).toEqual(beforeRejectedCommit);

    const deletedOwner = PptxDocument.create();
    deletedOwner.addSlide();
    deletedOwner.addSlide();
    const ownerGate = deferredSource();
    const pendingOwner = deletedOwner.addImage(0, ownerGate.source, {
      hyperlink: { slide: 2 },
    });
    deletedOwner.deleteSlide(0);
    deletedOwner.addSlide();
    const beforeOwnerReject = await sdkPackageSnapshot(deletedOwner);
    ownerGate.release();
    await expect(pendingOwner).rejects.toThrow(/owner slide was deleted/);
    expect(await sdkPackageSnapshot(deletedOwner)).toEqual(beforeOwnerReject);
  });

  it('adds detected raster image sources with immediate live model state', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    slide.addShape('rect', { name: 'Before source image' });
    const source = sdkPngHeader(640, 360);
    const expected = new Uint8Array(source);
    const options: AddImageSourceOptions = {
      name: 'Loaded & image',
      altText: 'Detected source',
      x: inches(1),
      y: inches(2),
      width: inches(3),
      height: inches(2),
      rotation: degrees(45),
      flipHorizontal: true,
      rounding: true,
      transparency: 25,
      shadow: {
        kind: 'outer',
        color: { kind: 'srgb', value: '123456' },
        opacity: 0.5,
        blur: 3,
        angle: 30,
        distance: 2,
        rotateWithShape: true,
      },
    };
    const pending: Promise<ImageModel> = document.addImage(0, source, options);
    source.fill(0);
    const image = await pending;

    expect(image).toBeInstanceOf(ImageModel);
    expect(slide.shapes[1]).toBe(image);
    expect(image.name).toBe('Loaded & image');
    expect(image.altText).toBe('Detected source');
    expect(image.rounding).toBe(true);
    expect(image.transparency).toBe(25);
    expect(image.shadow).toMatchObject({
      kind: 'outer',
      color: { kind: 'srgb', value: '123456' },
      rotateWithShape: true,
    });
    expect(image.transform).toEqual({
      x: inches(1),
      y: inches(2),
      width: inches(3),
      height: inches(2),
      rotation: degrees(45),
      flipHorizontal: true,
      flipVertical: false,
    });
    expect(image.sourcePartUri).toMatch(/\/ppt\/media\/image\d+\.png$/);
    expect(document.opcPackage.requirePart(image.sourcePartUri!)).toMatchObject({
      bytes: expected,
      contentType: 'image/png',
    });
    const slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml).toContain('name="Loaded &amp; image" descr="Detected source"');
    expect(slide.relationships.find(({ resolvedTarget }) => resolvedTarget === image.sourcePartUri))
      .toMatchObject({
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        targetMode: 'Internal',
      });
    expect(inspectRasterImage(expected)).toEqual({
      contentType: 'image/png',
      width: 640,
      height: 360,
    });

    const reopened = await PptxDocument.open(await document.write());
    const reopenedImage = reopened.slides[0]!.shapes[1] as ImageModel;
    expect(reopenedImage).toBeInstanceOf(ImageModel);
    expect(reopenedImage.name).toBe('Loaded & image');
    expect(reopenedImage.altText).toBe('Detected source');
    expect(reopenedImage.rounding).toBe(true);
    expect(reopenedImage.transparency).toBe(25);
    expect(reopenedImage.shadow).toMatchObject({
      kind: 'outer',
      color: { kind: 'srgb', value: '123456' },
      rotateWithShape: true,
    });
    expect(reopened.opcPackage.requirePart(reopenedImage.sourcePartUri!).bytes).toEqual(expected);
  });

  it('adds SVG sources with explicit and built-in PNG fallbacks through one live mutation', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const svg = sdkSvg(640, 360);
    const explicitFallback = sdkPngHeader(32, 18);
    const sizing: ImageSizing = {
      type: 'cover',
      width: inches(4),
      height: inches(3),
    };
    const source: ImageSource = svg;
    const pending = document.addImage(0, source, {
      contentType: 'image/svg+xml',
      fallback: explicitFallback,
      sizing,
      name: 'High-level SVG',
      altText: 'Vector with explicit fallback',
      x: inches(1),
      y: inches(2),
      rotation: degrees(15),
      flipHorizontal: true,
    });
    svg.fill(0);
    explicitFallback.fill(0);
    const image = await pending;

    expect(slide.shapes[0]).toBe(image);
    expect(image.isSvg).toBe(true);
    expect(image.name).toBe('High-level SVG');
    expect(image.transform).toEqual({
      x: inches(1),
      y: inches(2),
      width: inches(4),
      height: inches(3),
      rotation: degrees(15),
      flipHorizontal: true,
      flipVertical: false,
    });
    expect(image.sourceRectangle).toEqual({
      left: 12.5,
      top: 0,
      right: 12.5,
      bottom: 0,
    });
    expect(document.opcPackage.requirePart(image.fallbackPartUri!)).toMatchObject({
      contentType: 'image/png',
      bytes: sdkPngHeader(32, 18),
    });
    expect(document.opcPackage.requirePart(image.svgPartUri!)).toMatchObject({
      contentType: 'image/svg+xml',
      bytes: sdkSvg(640, 360),
    });

    const automatic = await document.addImage(0, sdkSvg(300, 150));
    expect(automatic.isSvg).toBe(true);
    expect(inspectRasterImage(
      document.opcPackage.requirePart(automatic.fallbackPartUri!).bytes,
    )).toEqual({ contentType: 'image/png', width: 1, height: 1 });
    expect(slide.shapes[1]).toBe(automatic);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedImage = reopened.slides[0]!.shapes[0] as ImageModel;
    expect(reopenedImage.isSvg).toBe(true);
    expect(reopenedImage.name).toBe('High-level SVG');
    expect(reopenedImage.sourceRectangle).toEqual(image.sourceRectangle);
  });

  it('imports arbitrary SVG and relationship prefixes without rewriting unrelated XML', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const image = slide.addSvgImage(sdkSvg(640, 360), sdkPngHeader(1, 1), {
      name: 'Alternate prefix SVG',
    });
    const part = document.opcPackage.requirePart(slide.partUri);
    const source = new TextDecoder().decode(part.bytes);
    const prefixed = source
      .replace(
        '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="',
        '<vector:svgBlip xmlns:vector="http://schemas.microsoft.com/office/drawing/2016/SVG/main" '
        + 'xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        + 'rel:embed="',
      )
      .replace(
        '</a:extLst>',
        '<a:ext uri="urn:keep"><keep:opaque xmlns:keep="urn:keep">UNCHANGED</keep:opaque>'
        + '</a:ext></a:extLst>',
      );
    expect(prefixed).not.toBe(source);
    document.opcPackage.setPart(slide.partUri, prefixed, part.contentType);

    const imported = await PptxDocument.open(await document.write());
    const importedSlide = imported.slides[0]!;
    const importedImage = importedSlide.shapes[0] as ImageModel;
    expect(importedImage.isSvg).toBe(true);
    expect(importedImage.fallbackPartUri).toBe(image.fallbackPartUri);
    expect(importedImage.svgPartUri).toBe(image.svgPartUri);
    const before = new Uint8Array(imported.opcPackage.requirePart(importedSlide.partUri).bytes);

    importedImage.replaceSvgData(sdkSvg(800, 600), sdkPngHeader(2, 2));

    expect(imported.opcPackage.requirePart(importedSlide.partUri).bytes).toEqual(before);
    const unchanged = new TextDecoder().decode(before);
    expect(unchanged).toContain('<vector:svgBlip');
    expect(unchanged).toContain('rel:embed=');
    expect(unchanged).toContain('<keep:opaque xmlns:keep="urn:keep">UNCHANGED</keep:opaque>');
    const reopened = await PptxDocument.open(await imported.write());
    const reopenedImage = reopened.slides[0]!.shapes[0] as ImageModel;
    expect(reopenedImage.isSvg).toBe(true);
    expect(reopened.opcPackage.requirePart(reopenedImage.svgPartUri!).bytes)
      .toEqual(sdkSvg(800, 600));
    expect(reopened.opcPackage.requirePart(reopenedImage.fallbackPartUri!).bytes)
      .toEqual(sdkPngHeader(2, 2));
  });

  it('rejects raster fallback without reading it and leaves failed SVG additions unchanged', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const snapshot = () => ({
      parts: document.opcPackage.parts.map(({ uri, contentType, bytes }) => ({
        uri,
        contentType,
        bytes: new Uint8Array(bytes),
      })),
      graph: document.opcPackage.graph,
      mutations: [...document.opcPackage.mutations],
      shapes: slide.shapes,
    });
    const before = snapshot();
    let fallbackReads = 0;
    const fallback: ImageSource = {
      async *[Symbol.asyncIterator]() {
        fallbackReads += 1;
        yield sdkPngHeader(1, 1);
      },
    };

    await expect(document.addImage(0, sdkPngHeader(10, 10), { fallback }))
      .rejects.toThrow(/fallback is only valid for SVG/i);
    expect(fallbackReads).toBe(0);
    expect(snapshot()).toEqual(before);

    await expect(document.addImage(0, sdkSvg(), {
      fallback: new TextEncoder().encode('GIF89a\x01\x00\x01\x00'),
    })).rejects.toThrow(/fallback.*PNG/i);
    expect(snapshot()).toEqual(before);

    const controller = new AbortController();
    const reason = new Error('abort explicit SVG fallback');
    const pending = document.addImage(0, sdkSvg(), {
      fallback: {
        async *[Symbol.asyncIterator]() {
          controller.abort(reason);
          yield sdkPngHeader(1, 1);
        },
      },
      signal: controller.signal,
    });
    await expect(pending).rejects.toBe(reason);
    expect(snapshot()).toEqual(before);
  });

  it('sizes every raster source form from intrinsic dimensions and round-trips all formats', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'pptx-image-sizing-'));
    const path = join(directory, 'landscape.png');
    const png = sdkPngHeader(1600, 900);
    await writeFile(path, png);
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(png.slice(0, 12));
        controller.enqueue(png.slice(12));
        controller.close();
      },
    });
    const iterable: RasterImageByteStream = {
      async *[Symbol.asyncIterator]() {
        yield png;
      },
    };
    const sources: RasterImageSource[] = [
      path,
      sdkPngDataUri(png),
      png,
      png.buffer.slice(0),
      new Blob([png], { type: 'text/plain' }),
      stream,
      iterable,
    ];
    const sizings = [
      {
        value: { type: 'cover' as const, width: inches(4), height: inches(3) },
        rectangle: { left: 12.5, top: 0, right: 12.5, bottom: 0 },
        raw: '<a:srcRect l="12500" t="0" r="12500" b="0"/>',
      },
      {
        value: { type: 'contain' as const, width: inches(4), height: inches(3) },
        rectangle: { left: 0, top: -16.667, right: 0, bottom: -16.667 },
        raw: '<a:srcRect l="0" t="-16667" r="0" b="-16667"/>',
      },
      {
        value: {
          type: 'crop' as const,
          width: inches(4),
          height: inches(3),
          source: { x: 400, y: 225, width: 800, height: 450 },
        },
        rectangle: { left: 25, top: 25, right: 25, bottom: 25 },
        raw: '<a:srcRect l="25000" t="25000" r="25000" b="25000"/>',
      },
    ];
    try {
      const document = PptxDocument.create();
      const slide = document.addSlide();
      for (const [index, source] of sources.entries()) {
        const sizing = sizings[index % sizings.length]!;
        const image = await document.addImage(0, source, {
          name: `Sized image ${index}`,
          altText: `Sizing mode ${sizing.value.type}`,
          x: inches(index),
          y: inches(index + 1),
          rotation: degrees(15),
          flipHorizontal: true,
          flipVertical: true,
          sizing: sizing.value,
        });
        expect(slide.shapes[index]).toBe(image);
        expect(image.transform).toEqual({
          x: inches(index),
          y: inches(index + 1),
          width: inches(4),
          height: inches(3),
          rotation: degrees(15),
          flipHorizontal: true,
          flipVertical: true,
        });
        expect(image.sourceRectangle).toEqual(sizing.rectangle);
      }

      const slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
      for (const sizing of sizings) expect(slideXml).toContain(sizing.raw);
      expect(slideXml).toContain('name="Sized image 0" descr="Sizing mode cover"');
      const reopened = await PptxDocument.open(await document.write());
      expect(reopened.slides[0]!.shapes).toHaveLength(sources.length);
      for (const [index, shape] of reopened.slides[0]!.shapes.entries()) {
        const image = shape as ImageModel;
        expect(image).toBeInstanceOf(ImageModel);
        expect(image.name).toBe(`Sized image ${index}`);
        expect(image.sourceRectangle).toEqual(sizings[index % sizings.length]!.rectangle);
        expect(image.transform).toMatchObject({ width: inches(4), height: inches(3) });
      }

      for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
        const formatted = PptxDocument.create({ format });
        formatted.addSlide();
        await formatted.addImage(0, png, {
          sizing: { type: 'cover', width: inches(4), height: inches(3) },
        });
        const formattedReopened = await PptxDocument.open(await formatted.write());
        const image = formattedReopened.slides[0]!.shapes[0] as ImageModel;
        expect(formattedReopened.format).toBe(format);
        expect(image.sourceRectangle).toEqual(sizings[0]!.rectangle);
        expect(image.transform).toMatchObject({ width: inches(4), height: inches(3) });
      }
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it('detaches sizing before awaiting raster source input', async () => {
    const document = PptxDocument.create();
    document.addSlide();
    const png = sdkPngHeader(1600, 900);
    const source = {
      async *[Symbol.asyncIterator]() {
        await Promise.resolve();
        yield png;
      },
    };
    const crop = { x: 400, y: 225, width: 800, height: 450 };
    const sizing = {
      type: 'crop' as const,
      width: inches(4),
      height: inches(3),
      source: crop,
    };
    const pending = document.addImage(0, source, { sizing });
    sizing.width = inches(1);
    crop.x = 0;

    const image = await pending;
    expect(image.transform).toMatchObject({ width: inches(4), height: inches(3) });
    expect(image.sourceRectangle).toEqual({ left: 25, top: 25, right: 25, bottom: 25 });
  });

  it('adds every portable in-memory source form and round-trips all formats', async () => {
    const png = sdkPngHeader(16, 9);
    const stream: RasterImageByteStream = new ReadableStream({
      start(controller) {
        controller.enqueue(png.slice(0, 12));
        controller.enqueue(png.slice(12));
        controller.close();
      },
    });
    const iterable: RasterImageByteStream = {
      async *[Symbol.asyncIterator]() {
        yield png;
      },
    };
    const sources: RasterImageSource[] = [
      png,
      png.buffer.slice(0),
      new Blob([png], { type: 'text/plain' }),
      stream,
      iterable,
      sdkPngDataUri(png),
    ];
    const document = PptxDocument.create();
    const slide = document.addSlide();
    for (const source of sources) await document.addImage(0, source);
    expect(slide.shapes).toHaveLength(sources.length);
    for (const [index, shape] of slide.shapes.entries()) {
      expect(shape).toBeInstanceOf(ImageModel);
      expect(shape.name).toBe(`Image ${index}`);
      expect(shape.transform).toEqual({
        x: 0,
        y: 0,
        width: inches(1),
        height: inches(1),
        rotation: 0,
        flipHorizontal: false,
        flipVertical: false,
      });
      expect(document.opcPackage.requirePart((shape as ImageModel).sourcePartUri!)).toMatchObject({
        bytes: png,
        contentType: 'image/png',
      });
    }

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      formatted.addSlide();
      const formattedImage = await formatted.addImage(0, sdkPngDataUri(png));
      expect(formattedImage).toBeInstanceOf(ImageModel);
      const reopened = await PptxDocument.open(await formatted.write());
      expect(reopened.format).toBe(format);
      expect(reopened.slides[0]!.shapes[0]).toBeInstanceOf(ImageModel);
      expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);
    }

    if (false) {
      const info: RasterImageInfo = inspectRasterImage(png);
      const typedSource: RasterImageSource = new Blob([png]);
      const typedOptions: AddImageSourceOptions = { contentType: 'image/png' };
      const percentageOptions: AddImageSourceOptions = {
        x: '10%',
        y: '20%',
        width: '30%',
        height: '40%',
      };
      const typedResult: Promise<ImageModel> = document.addImage(0, typedSource, typedOptions);
      const containResult: Promise<ImageModel> = document.addImage(0, typedSource, {
        sizing: { type: 'contain', width: inches(4), height: inches(3) },
      });
      const coverResult: Promise<ImageModel> = document.addImage(0, typedSource, {
        sizing: { type: 'cover', width: inches(4), height: inches(3) },
      });
      const cropResult: Promise<ImageModel> = document.addImage(0, typedSource, {
        sizing: {
          type: 'crop',
          width: inches(4),
          height: inches(3),
          source: { x: 0, y: 0, width: 16, height: 9 },
        },
      });
      const svgOptions: AddImageSourceOptions = { contentType: 'image/svg+xml' };
      const svgFallbackOptions: AddImageSourceOptions = {
        contentType: 'image/svg+xml',
        fallback: typedSource,
      };
      // @ts-expect-error sizing owns final width and cannot be combined with top-level width
      const conflictingWidth: AddImageSourceOptions = {
        width: inches(4),
        sizing: { type: 'cover', width: inches(4), height: inches(3) },
      };
      const directRectangle: AddImageSourceOptions = {
        // @ts-expect-error sourceRectangle belongs to the low-level model API
        sourceRectangle: { left: 0, top: 0, right: 0, bottom: 0 },
      };
      const malformedCrop: AddImageSourceOptions = {
        // @ts-expect-error crop sizing requires a source pixel region
        sizing: { type: 'crop', width: inches(4), height: inches(3) },
      };
      // @ts-expect-error plain objects are not raster image sources
      const invalidSource: RasterImageSource = { bytes: png };
      void [
        info,
        typedResult,
        percentageOptions,
        containResult,
        coverResult,
        cropResult,
        svgOptions,
        svgFallbackOptions,
        conflictingWidth,
        directRectangle,
        malformedCrop,
        invalidSource,
      ];
    }
  });

  it('creates and reopens raster and SVG percentage coordinates through the source loader', async () => {
    const document = PptxDocument.create({
      slideSize: { width: inches(10), height: inches(8) },
    });
    const slide = document.addSlide();
    const raster = await document.addImage(0, sdkPngHeader(16, 9), {
      x: '10%',
      y: '20%',
      width: '30%',
      height: '40%',
    });
    const vector = await document.addImage(0, sdkSvg(640, 360), {
      fallback: sdkPngHeader(1, 1),
      x: '12.5%',
      y: '25%',
      width: '37.5%',
      height: '50%',
    });

    const expected = [
      {
        x: inches(1),
        y: inches(1.6),
        width: inches(3),
        height: inches(3.2),
        rotation: 0,
        flipHorizontal: false,
        flipVertical: false,
      },
      {
        x: inches(1.25),
        y: inches(2),
        width: inches(3.75),
        height: inches(4),
        rotation: 0,
        flipHorizontal: false,
        flipVertical: false,
      },
    ];
    expect([raster.transform, vector.transform]).toEqual(expected);
    expect(vector.isSvg).toBe(true);

    const source = new TextDecoder().decode(
      document.opcPackage.requirePart(slide.partUri).bytes,
    );
    expect(source).toContain(
      '<a:off x="914400" y="1463040"/><a:ext cx="2743200" cy="2926080"/>',
    );
    expect(source).toContain(
      '<a:off x="1143000" y="1828800"/><a:ext cx="3429000" cy="3657600"/>',
    );

    await document.write({ compatibility: 'powerpoint-2010' });
    expect(document.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides[0]!.shapes.map(({ transform }) => transform)).toEqual(expected);
    expect((reopened.slides[0]!.shapes[1] as ImageModel).isSvg).toBe(true);
    expect(reopened.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);

    const before = await sdkPackageSnapshot(document);
    for (const options of [
      { x: '10%junk' },
      { width: '0%' },
      { height: '-10%' },
    ]) {
      await expect(document.addImage(
        0,
        sdkPngHeader(1, 1),
        options as AddImageSourceOptions,
      )).rejects.toBeInstanceOf(Error);
      expect(await sdkPackageSnapshot(document)).toEqual(before);
    }
  });

  it('rejects invalid document image additions without consuming unsafe input or mutating the package', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const before = () => ({
      parts: document.opcPackage.parts.map(({ uri, contentType, bytes }) => ({
        uri,
        contentType,
        bytes: new Uint8Array(bytes),
      })),
      graph: document.opcPackage.graph,
      mutations: [...document.opcPackage.mutations],
      relationships: slide.relationships,
      shapes: slide.shapes,
      shapeIds: slide.shapes.map(({ id }) => id),
    });
    const expected = before();
    const expectedZip = await document.write();
    let sourceReads = 0;
    const countedSource: RasterImageSource = {
      async *[Symbol.asyncIterator]() {
        sourceReads += 1;
        yield sdkPngHeader(1, 1);
      },
    };
    let accessorReads = 0;
    const accessor = Object.defineProperty({}, 'name', {
      enumerable: true,
      get() {
        accessorReads += 1;
        return 'unsafe';
      },
    });
    const sizingAccessor = Object.defineProperty(
      { type: 'cover', height: 1 },
      'width',
      {
        enumerable: true,
        get() {
          accessorReads += 1;
          return 1;
        },
      },
    );
    const cropSourceAccessor = Object.defineProperty(
      { x: 0, y: 0, height: 1 },
      'width',
      {
        enumerable: true,
        get() {
          accessorReads += 1;
          return 1;
        },
      },
    );
    const invalidPreIo: readonly (() => Promise<unknown>)[] = [
      () => document.addImage(1, countedSource),
      () => document.addImage(0, countedSource, null as never),
      () => document.addImage(0, countedSource, accessor as never),
      () => document.addImage(0, countedSource, Object.create({ name: 'inherited' })),
      () => document.addImage(0, countedSource, { unknown: true } as never),
      () => document.addImage(0, countedSource, { fallback: {} } as never),
      () => document.addImage(0, countedSource, { fallback: '' } as never),
      () => document.addImage(0, countedSource, { signal: {} } as never),
      () => document.addImage(0, countedSource, { sizing: [] } as never),
      () => document.addImage(0, countedSource, {
        sizing: { type: 'fit', width: 1, height: 1 },
      } as never),
      () => document.addImage(0, countedSource, { sizing: sizingAccessor } as never),
      () => document.addImage(0, countedSource, {
        sizing: {
          type: 'crop',
          width: 1,
          height: 1,
          source: cropSourceAccessor,
        },
      } as never),
      () => document.addImage(0, countedSource, {
        sizing: { type: 'cover', width: 1, height: 1 },
        width: 1,
      } as never),
      () => document.addImage(0, countedSource, {
        sizing: { type: 'crop', width: 1, height: 1, source: null },
      } as never),
      () => document.addImage(0, countedSource, {
        sourceRectangle: { left: 0, top: 0, right: 0, bottom: 0 },
      } as never),
      () => document.addImage(0, countedSource, { hyperlink: { slide: 99 } }),
    ];
    for (const invoke of invalidPreIo) {
      await expect(invoke()).rejects.toBeInstanceOf(Error);
      expect(before()).toEqual(expected);
      expect(await document.write()).toEqual(expectedZip);
    }
    expect(sourceReads).toBe(0);
    expect(accessorReads).toBe(0);

    const invalidAfterResolve: readonly (() => Promise<unknown>)[] = [
      () => document.addImage(0, sdkPngHeader(1, 1), { contentType: 'image/gif' }),
      () => document.addImage(0, sdkPngHeader(1, 1), { width: 0 } as never),
      () => document.addImage(0, Uint8Array.of(1, 2, 3)),
      () => document.addImage(0, 'data:image/png;base64,AAAA'),
      () => document.addImage(0, sdkPngHeader(100, 80), {
        sizing: {
          type: 'crop',
          width: 1,
          height: 1,
          source: { x: 90, y: 0, width: 11, height: 1 },
        },
      }),
      () => document.addImage(0, sdkPngHeader(0xffff_ffff, 1), {
        sizing: { type: 'contain', width: 1, height: Number.MAX_SAFE_INTEGER },
      }),
    ];
    for (const invoke of invalidAfterResolve) {
      await expect(invoke()).rejects.toBeInstanceOf(Error);
      expect(before()).toEqual(expected);
      expect(await document.write()).toEqual(expectedZip);
    }
  }, 10_000);

  it('creates preset shapes with deterministic defaults, transforms, order, and identity', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const relationships = slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }));

    const rectangle = slide.addShape('rect');
    const line = slide.addShape('lineInv', {
      name: 'A & <Line>',
      x: inches(2),
      y: inches(3),
      width: inches(4),
      height: inches(0.5),
      rotation: degrees(45),
      flipHorizontal: true,
      flipVertical: true,
    });

    expect([rectangle.id, line.id]).toEqual([2, 3]);
    expect([rectangle.kind, line.kind]).toEqual(['shape', 'shape']);
    expect(rectangle).toBeInstanceOf(ShapeModel);
    expect(rectangle.name).toBe('Shape 2');
    expect(rectangle.transform).toEqual({
      x: inches(1),
      y: inches(1),
      width: inches(1),
      height: inches(1),
      rotation: 0,
      flipHorizontal: false,
      flipVertical: false,
    });
    expect(line.name).toBe('A & <Line>');
    expect(line.transform).toEqual({
      x: inches(2),
      y: inches(3),
      width: inches(4),
      height: inches(0.5),
      rotation: degrees(45),
      flipHorizontal: true,
      flipVertical: true,
    });
    expect(slide.shapes).toEqual([rectangle, line]);
    expect(slide.shapes[0]).toBe(rectangle);
    expect(slide.shapes[1]).toBe(line);
    expect(slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }))).toEqual(relationships);

    const xml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(xml).toContain('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/>');
    expect(xml).toContain('<a:xfrm rot="2700000" flipH="1" flipV="1">');
    expect(xml).toContain('name="A &amp; &lt;Line&gt;"');
    expect(xml).not.toContain('<p:txBody>');
  });

  it('creates and reopens shape and text percentage coordinates against the current slide size', async () => {
    const document = PptxDocument.create({
      slideSize: { width: inches(10), height: inches(8) },
    });
    const slide = document.addSlide();
    const geometry: CustomGeometry = {
      paths: [{
        width: 100,
        height: 100,
        commands: [
          { kind: 'moveTo', point: { x: 0, y: 0 } },
          { kind: 'lineTo', point: { x: 100, y: 100 } },
        ],
      }],
    };
    const shape = slide.addShape('rect', {
      x: '10%',
      y: '20%',
      width: '30%',
      height: '40%',
    });
    const custom = slide.addCustomShape(geometry, {
      x: '-10%',
      y: '125%',
      width: '20%',
      height: '25%',
    });
    const text = slide.addText('Percentage text', {
      x: '12.5%',
      y: '25%',
      width: '37.5%',
      height: '50%',
    });
    const richText = slide.addRichText([{
      runs: [{ text: 'Percentage rich text' }],
    }], {
      x: inches(1),
      y: '50%',
      width: '50%',
      height: inches(1),
    });

    const expectedTransforms = [
      {
        x: inches(1),
        y: inches(1.6),
        width: inches(3),
        height: inches(3.2),
        rotation: 0,
        flipHorizontal: false,
        flipVertical: false,
      },
      {
        x: inches(-1),
        y: inches(10),
        width: inches(2),
        height: inches(2),
        rotation: 0,
        flipHorizontal: false,
        flipVertical: false,
      },
      {
        x: inches(1.25),
        y: inches(2),
        width: inches(3.75),
        height: inches(4),
        rotation: 0,
        flipHorizontal: false,
        flipVertical: false,
      },
      {
        x: inches(1),
        y: inches(4),
        width: inches(5),
        height: inches(1),
        rotation: 0,
        flipHorizontal: false,
        flipVertical: false,
      },
    ];
    expect([shape.transform, custom.transform, text.transform, richText.transform])
      .toEqual(expectedTransforms);

    const xml = new TextDecoder().decode(
      document.opcPackage.requirePart(slide.partUri).bytes,
    );
    for (const [x, y, width, height] of [
      [914_400, 1_463_040, 2_743_200, 2_926_080],
      [-914_400, 9_144_000, 1_828_800, 1_828_800],
      [1_143_000, 1_828_800, 3_429_000, 3_657_600],
      [914_400, 3_657_600, 4_572_000, 914_400],
    ]) {
      expect(xml).toContain(
        `<a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/>`,
      );
    }

    const beforeInvalid = await sdkPackageSnapshot(document);
    let accessorReads = 0;
    const accessor = Object.defineProperty({
      y: '10%',
      width: '10%',
      height: '10%',
    }, 'x', {
      enumerable: true,
      get() {
        accessorReads += 1;
        return '10%';
      },
    });
    for (const invoke of [
      () => slide.addShape('rect', { width: '0%' }),
      () => slide.addCustomShape(geometry, { height: '-1%' }),
      () => slide.addText('Malformed', { x: '10%oops' } as never),
      () => slide.addText('Accessor', accessor as never),
    ]) {
      expect(invoke).toThrow();
      expect(await sdkPackageSnapshot(document)).toEqual(beforeInvalid);
    }
    expect(accessorReads).toBe(0);

    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);
    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides[0]!.shapes.map(({ transform }) => transform))
      .toEqual(expectedTransforms);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);
  });

  it('creates and reopens styled custom geometry shapes through the public SDK', async () => {
    const document = PptxDocument.create();
    const source = document.addSlide();
    const target = document.addSlide();
    const geometry: CustomGeometry = {
      paths: [{
        width: inches(4),
        height: inches(3),
        commands: [
          { kind: 'moveTo', point: { x: 0, y: 0 } },
          { kind: 'lineTo', point: { x: inches(4), y: 0 } },
          { kind: 'lineTo', point: { x: inches(2), y: inches(3) } },
          { kind: 'close' },
        ],
      }],
    };
    const mutable = structuredClone(geometry) as {
      paths: Array<{
        width: number;
        height: number;
        commands: Array<CustomGeometry['paths'][number]['commands'][number]>;
      }>;
    };
    const custom = source.addCustomShape(mutable, {
      name: 'SDK custom triangle',
      x: inches(1),
      y: inches(1),
      width: inches(4),
      height: inches(3),
      rotation: degrees(15),
      flipVertical: true,
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
      line: { kind: 'line', color: { kind: 'srgb', value: '112233' }, width: 2 },
      arrows: { end: 'triangle' },
      shadow: { kind: 'outer' },
      hyperlink: { slide: 2, tooltip: 'Target' },
    });
    const multi = source.addCustomShape({
      paths: [
        {
          width: 100,
          height: 200,
          fill: 'none',
          stroke: false,
          commands: [{ kind: 'moveTo', point: { x: 0, y: 0 } }],
        },
        { width: 300, height: 400, commands: [] },
      ],
    }, { name: 'SDK multiple paths' });
    mutable.paths[0]!.width = 1;
    mutable.paths[0]!.commands.splice(0);

    expect([custom.id, multi.id]).toEqual([2, 3]);
    expect(source.shapes).toEqual([custom, multi]);
    expect(source.shapes[0]).toBe(custom);
    expect(custom).toBeInstanceOf(ShapeModel);
    expect(custom.kind).toBe('shape');
    expect(custom.presetType).toBeUndefined();
    expect(custom.name).toBe('SDK custom triangle');
    expect(custom.transform).toEqual({
      x: inches(1),
      y: inches(1),
      width: inches(4),
      height: inches(3),
      rotation: degrees(15),
      flipHorizontal: false,
      flipVertical: true,
    });
    expect(custom.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
    });
    expect(custom.line).toMatchObject({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      width: 2,
    });
    expect(custom.arrows).toEqual({ end: 'triangle' });
    expect(custom.shadow).toMatchObject({ kind: 'outer' });
    expect(custom.hyperlink).toEqual({ slide: 2, tooltip: 'Target' });

    const sourceXml = new TextDecoder().decode(
      document.opcPackage.requirePart(source.partUri).bytes,
    );
    expect(sourceXml).toContain('<a:path w="3657600" h="2743200">');
    expect(sourceXml).toContain('<a:lnTo><a:pt x="3657600" y="0"/></a:lnTo>');
    expect(sourceXml).toContain('<a:path w="300" h="400"></a:path>');
    expect(sourceXml.match(/<a:custGeom>/g)).toHaveLength(2);
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    const duplicate = document.duplicateSlide(document.slides.indexOf(source));
    document.moveSlide(document.slides.indexOf(duplicate), 0);
    expect(new TextDecoder().decode(
      document.opcPackage.requirePart(duplicate.partUri).bytes,
    ).match(/<a:custGeom>/g)).toHaveLength(2);
    document.deleteSlide(document.slides.indexOf(duplicate));
    const reopened = await PptxDocument.open(await document.write());
    const reopenedSource = reopened.slides.find(({ partUri }) => partUri === source.partUri)!;
    const reopenedShape = reopenedSource.shapes[0] as ShapeModel;
    expect(reopenedShape.name).toBe('SDK custom triangle');
    expect(reopenedShape.presetType).toBeUndefined();
    expect(reopenedShape.hyperlink).toEqual({ slide: 2, tooltip: 'Target' });
    expect(new TextDecoder().decode(
      reopened.opcPackage.requirePart(reopenedSource.partUri).bytes,
    ).match(/<a:custGeom>/g)).toHaveLength(2);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      formatted.addSlide().addCustomShape(geometry, { name: `Custom ${format}` });
      const formattedReopened = await PptxDocument.open(await formatted.write());
      expect(formattedReopened.format).toBe(format);
      expect(formattedReopened.slides[0]!.shapes[0]!.name).toBe(`Custom ${format}`);
      expect(new TextDecoder().decode(
        formattedReopened.opcPackage.requirePart(formattedReopened.slides[0]!.partUri).bytes,
      )).toContain('<a:custGeom>');
    }

    if (false) {
      // @ts-expect-error custom shape options exclude preset-only adjustments
      source.addCustomShape(geometry, { adjustments: [] });
      source.addCustomShape({
        paths: [{
          width: 1,
          height: 1,
          commands: [{
            kind: 'moveTo',
            // @ts-expect-error custom geometry coordinates are number or string values
            point: { x: false, y: 2 },
          }],
        }],
      });
      // Keep target live for the compile-time branch.
      void target;
    }
  });

  it('exposes pure and live custom geometry evaluation through the public SDK', async () => {
    const geometry: CustomGeometry = {
      adjustments: [{ name: 'adj', formula: { operator: 'val', operands: [50_000] } }],
      guides: [
        { name: 'x1', formula: { operator: '*/', operands: ['w', 'adj', 100_000] } },
        { name: 'y1', formula: { operator: '*/', operands: ['h', 'adj', 100_000] } },
      ],
      handles: [{
        kind: 'xy',
        position: { x: 'x1', y: 'y1' },
        xGuide: 'adj',
        minX: 'l',
        maxX: 'r',
      }],
      connectionSites: [{ angle: 'cd4', position: { x: 'x1', y: 'y1' } }],
      textRectangle: { left: 'x1', top: 'y1', right: 'r', bottom: 'b' },
      paths: [{
        width: 100_000,
        height: 100_000,
        commands: [
          { kind: 'moveTo', point: { x: 'x1', y: 'y1' } },
          { kind: 'lineTo', point: { x: 'r', y: 'b' } },
        ],
      }],
    };
    const context: CustomGeometryEvaluationContext = {
      width: inches(2),
      height: inches(1),
    };
    const pure: EvaluatedCustomGeometry = evaluateCustomGeometry(geometry, context);
    expect(typeof evaluateCustomGeometry).toBe('function');
    expect(pure.context).toEqual(context);
    expect(pure.guides).toEqual([
      { name: 'x1', value: inches(1) },
      { name: 'y1', value: inches(0.5) },
    ]);
    expect(pure.textRectangle).toEqual({
      left: inches(1),
      top: inches(0.5),
      right: inches(2),
      bottom: inches(1),
    });
    expect(Object.isFrozen(pure)).toBe(true);
    expect(Object.isFrozen(pure.context)).toBe(true);
    expect(Object.isFrozen(pure.handles?.[0]?.position)).toBe(true);
    expect(Object.isFrozen(pure.paths[0]?.commands[0])).toBe(true);

    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addCustomShape(geometry, {
      width: inches(2),
      height: inches(1),
    });
    const bytes = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const journal = [...document.opcPackage.mutations];
    const live = shape.evaluateCustomGeometry();
    expect(live).toEqual(pure);
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(bytes);
    expect(document.opcPackage.mutations).toEqual(journal);

    shape.setTransform({ width: inches(3), height: inches(1.5) });
    expect(shape.evaluateCustomGeometry()?.guides).toEqual([
      { name: 'x1', value: inches(1.5) },
      { name: 'y1', value: inches(0.75) },
    ]);
    const reopened = await PptxDocument.open(await document.write());
    const reopenedShape = reopened.slides[0]!.shapes[0] as ShapeModel;
    expect(reopenedShape.evaluateCustomGeometry()).toEqual(shape.evaluateCustomGeometry());

    const unknownGeometry: CustomGeometry = {
      paths: [{
        width: 1,
        height: 1,
        commands: [{ kind: 'moveTo', point: { x: 'missing', y: 0 } }],
      }],
    };
    const unknownSlide = document.addSlide();
    const unknownShape = unknownSlide.addCustomShape(unknownGeometry);
    expect(unknownShape.customGeometry).toEqual(unknownGeometry);
    try {
      unknownShape.evaluateCustomGeometry();
      throw new Error('Expected custom geometry evaluation to fail');
    } catch (error) {
      expect(error).toBeInstanceOf(CustomGeometryEvaluationError);
      expect(error).toMatchObject({ code: 'unknown-token', token: 'missing' });
    }

    if (false) {
      // @ts-expect-error live evaluator reads its context from the shape transform
      shape.evaluateCustomGeometry(context);
    }
  });

  it('edits and converts custom geometry through the public SDK lifecycle', async () => {
    const geometry: CustomGeometry = {
      paths: [{
        width: inches(4),
        height: inches(3),
        commands: [
          { kind: 'moveTo', point: { x: 0, y: 0 } },
          { kind: 'lineTo', point: { x: inches(4), y: 0 } },
          { kind: 'lineTo', point: { x: inches(2), y: inches(3) } },
          { kind: 'close' },
        ],
      }],
    };
    const replacement: CustomGeometry = {
      paths: [
        {
          width: 100,
          height: 200,
          fill: 'none',
          stroke: false,
          commands: [
            { kind: 'moveTo', point: { x: 1, y: 2 } },
            {
              kind: 'quadraticBezierTo',
              control: { x: 3, y: 4 },
              end: { x: 5, y: 6 },
            },
          ],
        },
        { width: 300, height: 400, extrusionOk: false, commands: [] },
      ],
    };
    const document = PptxDocument.create();
    const source = document.addSlide();
    const shape = source.addCustomShape(geometry, {
      name: 'Editable SDK custom',
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
      line: { kind: 'line', color: { kind: 'srgb', value: '112233' }, width: 2 },
      shadow: { kind: 'outer' },
      hyperlink: { url: 'https://example.com/sdk-custom' },
    });
    const first: CustomGeometry | undefined = shape.customGeometry;
    expect(first).toEqual(geometry);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(first?.paths[0]?.commands)).toBe(true);
    expect(shape.customGeometry).not.toBe(first);

    const before = document.opcPackage.requirePart(source.partUri).bytes.slice();
    const journal = [...document.opcPackage.mutations];
    shape.customGeometry = structuredClone(geometry);
    expect(document.opcPackage.requirePart(source.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);

    const identity = shape;
    shape.customGeometry = replacement;
    expect(shape).toBe(identity);
    expect(source.shapes[0]).toBe(identity);
    expect(shape.customGeometry).toEqual(replacement);
    expect(shape.name).toBe('Editable SDK custom');
    expect(shape.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
    });
    expect(shape.hyperlink).toEqual({ url: 'https://example.com/sdk-custom' });

    const duplicate = document.duplicateSlide(document.slides.indexOf(source));
    const duplicateShape = duplicate.shapes[0] as ShapeModel;
    shape.presetType = 'ellipse';
    expect(shape.presetType).toBe('ellipse');
    expect(shape.customGeometry).toBeUndefined();
    expect(duplicateShape.presetType).toBeUndefined();
    expect(duplicateShape.customGeometry).toEqual(replacement);

    const rollbackBytes = document.opcPackage.requirePart(duplicate.partUri).bytes.slice();
    const rollbackJournal = [...document.opcPackage.mutations];
    expect(() => document.transaction(() => {
      duplicateShape.presetType = 'star5';
      expect(duplicateShape.customGeometry).toBeUndefined();
      throw new Error('restore SDK geometry conversion');
    })).toThrow('restore SDK geometry conversion');
    expect(document.opcPackage.requirePart(duplicate.partUri).bytes).toEqual(rollbackBytes);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(duplicateShape.customGeometry).toEqual(replacement);

    duplicateShape.presetType = 'star5';
    duplicateShape.customGeometry = geometry;
    expect(duplicateShape.presetType).toBeUndefined();
    expect(duplicateShape.customGeometry).toEqual(geometry);
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedSource = reopened.slides.find(({ partUri }) => partUri === source.partUri)!;
    const reopenedDuplicate = reopened.slides.find(({ partUri }) => partUri === duplicate.partUri)!;
    const reopenedPreset = reopenedSource.shapes[0] as ShapeModel;
    const reopenedCustom = reopenedDuplicate.shapes[0] as ShapeModel;
    expect(reopenedPreset.presetType).toBe('ellipse');
    expect(reopenedPreset.customGeometry).toBeUndefined();
    expect(reopenedCustom.presetType).toBeUndefined();
    expect(reopenedCustom.customGeometry).toEqual(geometry);
    expect(reopenedCustom.name).toBe('Editable SDK custom');
    expect(reopenedCustom.hyperlink).toEqual({ url: 'https://example.com/sdk-custom' });

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      const formattedShape = formatted.addSlide().addCustomShape(geometry);
      formattedShape.customGeometry = replacement;
      formattedShape.presetType = 'diamond';
      formattedShape.customGeometry = geometry;
      const formattedReopened = await PptxDocument.open(await formatted.write());
      const reopenedShape = formattedReopened.slides[0]!.shapes[0] as ShapeModel;
      expect(formattedReopened.format).toBe(format);
      expect(reopenedShape.presetType).toBeUndefined();
      expect(reopenedShape.customGeometry).toEqual(geometry);
    }

    if (false) {
      // @ts-expect-error custom geometry cannot be cleared from a shape
      shape.customGeometry = undefined;
    }
  });

  it('creates and edits custom geometry guide formulas through the public SDK', async () => {
    const geometry: CustomGeometry = {
      adjustments: [{
        name: 'adj1',
        formula: { operator: 'val', operands: [25_000] },
      }],
      guides: [
        { name: 'x1', formula: { operator: '*/', operands: ['w', 'adj1', 100_000] } },
        { name: 'y1', formula: { operator: '+-', operands: ['h', 0, 'x1'] } },
        { name: 'a1', formula: { operator: 'at2', operands: ['y1', 'x1'] } },
      ],
      paths: [{
        width: 100_000,
        height: 100_000,
        commands: [
          { kind: 'moveTo', point: { x: 'x1', y: 0 } },
          { kind: 'lineTo', point: { x: 'r', y: 'y1' } },
          {
            kind: 'arcTo',
            widthRadius: 'x1',
            heightRadius: 'hd2',
            startAngle: 'a1',
            sweepAngle: 'cd2',
          },
        ],
      }],
    };
    const replacement: CustomGeometry = {
      adjustments: [{
        name: 'adj1',
        formula: { operator: 'val', operands: [50_000] },
      }],
      guides: [{
        name: 'x1',
        formula: { operator: 'pin', operands: [0, 'adj1', 100_000] },
      }],
      paths: [{
        width: 100_000,
        height: 100_000,
        fill: 'none',
        commands: [
          { kind: 'moveTo', point: { x: 'x1', y: 't' } },
          { kind: 'lineTo', point: { x: 'r', y: 'b' } },
        ],
      }],
    };

    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addCustomShape(geometry, { name: 'SDK formula geometry' });
    expect(shape.customGeometry).toEqual(geometry);
    expect(Object.isFrozen(shape.customGeometry?.adjustments?.[0]?.formula.operands)).toBe(true);
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);
    const xml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(xml).toContain('<a:gd name="x1" fmla="*/ w adj1 100000"/>');
    expect(xml).toContain('<a:arcTo wR="x1" hR="hd2" stAng="a1" swAng="cd2"/>');

    const reopened = await PptxDocument.open(await document.write());
    const reopenedShape = reopened.slides[0]!.shapes[0] as ShapeModel;
    expect(reopenedShape.customGeometry).toEqual(geometry);
    reopenedShape.customGeometry = replacement;
    expect(reopenedShape.customGeometry).toEqual(replacement);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    const edited = await PptxDocument.open(await reopened.write());
    const editedShape = edited.slides[0]!.shapes[0] as ShapeModel;
    expect(editedShape.name).toBe('SDK formula geometry');
    expect(editedShape.customGeometry).toEqual(replacement);
  });

  it('creates, reopens, and edits custom geometry adjustment handles through the public SDK', async () => {
    const geometry: CustomGeometry = {
      adjustments: [
        { name: 'adjX', formula: { operator: 'val', operands: [25_000] } },
        { name: 'adjY', formula: { operator: 'val', operands: [50_000] } },
        { name: 'adjR', formula: { operator: 'val', operands: [30_000] } },
        { name: 'adjAng', formula: { operator: 'val', operands: [5_400_000] } },
      ],
      guides: [
        { name: 'x1', formula: { operator: '*/', operands: ['w', 'adjR', 100_000] } },
        { name: 'y1', formula: { operator: '*/', operands: ['h', 'adjR', 100_000] } },
      ],
      handles: [
        {
          kind: 'xy',
          position: { x: 'adjX', y: 'adjY' },
          xGuide: 'adjX',
          minX: 0,
          maxX: 100_000,
          yGuide: 'adjY',
          minY: 't',
          maxY: 'b',
        },
        {
          kind: 'polar',
          position: { x: 'x1', y: 'y1' },
          radiusGuide: 'adjR',
          minRadius: 0,
          maxRadius: 'ss',
          angleGuide: 'adjAng',
          minAngle: 0,
          maxAngle: 'cd',
        },
      ],
      paths: [{
        width: 100_000,
        height: 100_000,
        commands: [
          { kind: 'moveTo', point: { x: 'adjX', y: 0 } },
          { kind: 'lineTo', point: { x: 'r', y: 'adjY' } },
          { kind: 'close' },
        ],
      }],
    };
    const replacement: CustomGeometry = {
      ...geometry,
      handles: [
        {
          kind: 'polar',
          position: { x: 'hc', y: 'vc' },
          radiusGuide: 'adjR',
          minRadius: 1,
          maxRadius: 'ss',
          angleGuide: 'adjAng',
          minAngle: 0,
          maxAngle: '3cd4',
        },
        {
          kind: 'xy',
          position: { x: 'x1', y: 'adjY' },
          xGuide: 'adjX',
          minX: 0,
          maxX: 90_000,
          yGuide: 'adjY',
          minY: 't',
          maxY: 'b',
        },
      ],
    };

    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addCustomShape(geometry, { name: 'SDK handle geometry' });
    expect(shape.customGeometry).toEqual(geometry);
    expect(Object.isFrozen(shape.customGeometry?.handles)).toBe(true);
    expect(shape.customGeometry?.handles?.every((handle) =>
      Object.isFrozen(handle) && Object.isFrozen(handle.position))).toBe(true);
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedShape = reopened.slides[0]!.shapes[0] as ShapeModel;
    expect(reopenedShape.customGeometry).toEqual(geometry);
    reopenedShape.customGeometry = replacement;
    expect(reopenedShape.customGeometry).toEqual(replacement);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    const edited = await PptxDocument.open(await reopened.write());
    const editedSlide = edited.slides[0]!;
    const editedShape = editedSlide.shapes[0] as ShapeModel;
    const xml = new TextDecoder().decode(
      edited.opcPackage.requirePart(editedSlide.partUri).bytes,
    );
    expect(editedShape.name).toBe('SDK handle geometry');
    expect(editedShape.customGeometry).toEqual(replacement);
    expect(xml.indexOf('<a:ahPolar')).toBeLessThan(xml.indexOf('<a:ahXY'));
    expect(validatePackage(edited.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);
  });

  it('creates, edits, converts, and reopens custom geometry connection sites through the SDK', async () => {
    const geometry: CustomGeometry = {
      adjustments: [{
        name: 'adjAng',
        formula: { operator: 'val', operands: [5_400_000] },
      }],
      connectionSites: [
        { angle: 0, position: { x: 'hc', y: 't' } },
        { angle: 'adjAng', position: { x: 'r', y: 'vc' } },
        { angle: -5_400_000, position: { x: 25_000, y: 100_000 } },
        { angle: 0, position: { x: 'hc', y: 't' } },
      ],
      paths: [{
        width: 100_000,
        height: 100_000,
        commands: [
          { kind: 'moveTo', point: { x: 0, y: 0 } },
          { kind: 'lineTo', point: { x: 'r', y: 'b' } },
        ],
      }],
    };
    const replacement: CustomGeometry = {
      ...geometry,
      connectionSites: [
        { angle: -5_400_000, position: { x: 25_000, y: 100_000 } },
        { angle: 0, position: { x: 'hc', y: 't' } },
        { angle: 'adjAng', position: { x: 'l', y: 'vc' } },
        { angle: 0, position: { x: 'hc', y: 't' } },
      ],
    };
    const mutable = structuredClone(geometry) as unknown as {
      connectionSites: Array<{
        angle: string | number;
        position: { x: string | number; y: string | number };
      }>;
    };

    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addCustomShape(mutable as unknown as CustomGeometry, {
      name: 'SDK connection geometry',
    });
    mutable.connectionSites[0]!.angle = 1;
    mutable.connectionSites[0]!.position.x = 'changed';
    mutable.connectionSites.reverse();
    expect(shape.customGeometry).toEqual(geometry);
    expect(Object.isFrozen(shape.customGeometry?.connectionSites)).toBe(true);
    expect(shape.customGeometry?.connectionSites?.every((site) =>
      Object.isFrozen(site) && Object.isFrozen(site.position))).toBe(true);

    const before = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const journal = [...document.opcPackage.mutations];
    shape.customGeometry = structuredClone(geometry);
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);

    shape.customGeometry = replacement;
    expect(shape.customGeometry).toEqual(replacement);
    shape.presetType = 'diamond';
    expect(shape.presetType).toBe('diamond');
    expect(shape.customGeometry).toBeUndefined();
    shape.customGeometry = replacement;
    expect(shape.presetType).toBeUndefined();
    expect(shape.customGeometry).toEqual(replacement);
    expect(shape.name).toBe('SDK connection geometry');
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedShape = reopened.slides[0]!.shapes[0] as ShapeModel;
    expect(reopenedShape.name).toBe('SDK connection geometry');
    expect(reopenedShape.customGeometry).toEqual(replacement);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);
  });

  it('creates, edits, resets, converts, and reopens custom geometry text rectangles through the SDK', async () => {
    const geometry: CustomGeometry = {
      guides: [
        { name: 'textLeft', formula: { operator: 'val', operands: [20_000] } },
        { name: 'textRight', formula: { operator: 'val', operands: [80_000] } },
      ],
      connectionSites: [{ angle: 0, position: { x: 'hc', y: 't' } }],
      textRectangle: {
        left: 'textLeft',
        top: 10_000,
        right: 'textRight',
        bottom: 90_000,
      },
      paths: [{
        width: 100_000,
        height: 100_000,
        commands: [
          { kind: 'moveTo', point: { x: 0, y: 0 } },
          { kind: 'lineTo', point: { x: 'r', y: 'b' } },
        ],
      }],
    };
    const replacement: CustomGeometry = {
      ...geometry,
      textRectangle: {
        left: 0,
        top: 't',
        right: 85_000,
        bottom: 'b',
      },
    };
    const mutable = structuredClone(geometry) as unknown as {
      textRectangle: {
        left: string | number;
        top: string | number;
        right: string | number;
        bottom: string | number;
      };
    };

    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addCustomShape(mutable as unknown as CustomGeometry, {
      name: 'SDK text rectangle geometry',
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
      line: { kind: 'line', color: { kind: 'srgb', value: '112233' }, width: 2 },
      hyperlink: { url: 'https://example.com/sdk-text-rectangle' },
    });
    mutable.textRectangle.left = 'changed';
    mutable.textRectangle.top = 1;
    mutable.textRectangle.right = 2;
    mutable.textRectangle.bottom = 3;

    const first = shape.customGeometry;
    expect(first).toEqual(geometry);
    expect(shape.customGeometry).not.toBe(first);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(first?.textRectangle)).toBe(true);
    const before = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const journal = [...document.opcPackage.mutations];
    shape.customGeometry = structuredClone(geometry);
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);

    const identity = shape;
    shape.customGeometry = replacement;
    expect(shape).toBe(identity);
    expect(slide.shapes[0]).toBe(identity);
    expect(shape.customGeometry).toEqual(replacement);
    expect(shape.name).toBe('SDK text rectangle geometry');
    expect(shape.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
    });
    expect(shape.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      width: 2,
      dash: 'solid',
    });
    expect(shape.hyperlink).toEqual({ url: 'https://example.com/sdk-text-rectangle' });

    const { textRectangle: _textRectangle, ...defaultTextRectangleGeometry } = replacement;
    shape.customGeometry = defaultTextRectangleGeometry;
    expect(Object.hasOwn(shape.customGeometry!, 'textRectangle')).toBe(false);
    expect(new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes))
      .toContain('<a:rect l="l" t="t" r="r" b="b"/>');

    shape.customGeometry = replacement;
    shape.presetType = 'diamond';
    expect(shape.presetType).toBe('diamond');
    expect(shape.customGeometry).toBeUndefined();
    shape.customGeometry = replacement;
    expect(shape.presetType).toBeUndefined();
    expect(shape.customGeometry).toEqual(replacement);
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedShape = reopened.slides[0]!.shapes[0] as ShapeModel;
    expect(reopenedShape.name).toBe('SDK text rectangle geometry');
    expect(reopenedShape.customGeometry).toEqual(replacement);
    expect(Object.isFrozen(reopenedShape.customGeometry?.textRectangle)).toBe(true);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);
  });

  it('creates preset shape hyperlinks through the public SDK type and runtime surface', () => {
    const document = PptxDocument.create();
    const first = document.addSlide();
    const second = document.addSlide();
    const websiteTarget: Hyperlink = {
      url: 'https://example.com?a=1&b=2',
      tooltip: 'Visit & learn',
    };
    const slideTarget: Hyperlink = { slide: 2, tooltip: '' };

    const website = first.addShape('rect', { hyperlink: websiteTarget });
    const next = first.addShape('actionButtonForwardNext', { hyperlink: slideTarget });
    const self = first.addShape('actionButtonHome', { hyperlink: { slide: 1 } });

    expect(first.shapes).toEqual([website, next, self]);
    expect(first.shapes[0]).toBe(website);
    expect(first.relationships.find(({ type }) => type.endsWith('/hyperlink'))).toMatchObject({
      target: 'https://example.com?a=1&b=2',
      targetMode: 'External',
    });
    expect(first.relationships.filter(({ type }) => type.endsWith('/slide'))).toEqual([
      expect.objectContaining({ targetMode: 'Internal', resolvedTarget: second.partUri }),
      expect.objectContaining({ targetMode: 'Internal', resolvedTarget: first.partUri }),
    ]);
    const xml = new TextDecoder().decode(document.opcPackage.requirePart(first.partUri).bytes);
    expect(xml).toContain('tooltip="Visit &amp; learn"');
    expect(xml).toContain('tooltip="" action="ppaction://hlinksldjump"');
    expect(xml.match(/<a:hlinkClick/g)).toHaveLength(3);
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    if (false) {
      // @ts-expect-error shape hyperlink requires exactly one target
      first.addShape('rect', { hyperlink: {} });
      // @ts-expect-error shape hyperlink target branches are mutually exclusive
      first.addShape('rect', { hyperlink: { url: 'https://example.com', slide: 2 } });
      // @ts-expect-error shape hyperlink URL must be a string
      first.addShape('rect', { hyperlink: { url: 42 } });
      // @ts-expect-error shape hyperlink slide must be numeric
      first.addShape('rect', { hyperlink: { slide: '2' } });
      // @ts-expect-error shape hyperlink public value has no relationship ID escape hatch
      first.addShape('rect', { hyperlink: { url: 'https://example.com', _rId: 'rId9' } });
      // @ts-expect-error shape hyperlink tooltip must be a string
      first.addShape('rect', { hyperlink: { slide: 2, tooltip: 7 } });
    }
  });

  it('preserves the shape hyperlink lifecycle through duplicate, move, delete, and all formats', async () => {
    const document = PptxDocument.create();
    const source = document.addSlide();
    const target = document.addSlide();
    const external = source.addShape('rect', {
      hyperlink: { url: 'https://example.com', tooltip: 'Visit' },
    });
    const internal = source.addShape('actionButtonForwardNext', {
      hyperlink: { slide: 2 },
    });
    const self = source.addShape('actionButtonHome', {
      hyperlink: { slide: 1, tooltip: '' },
    });
    const duplicate = document.duplicateSlide(document.slides.indexOf(source));
    const [duplicateExternal, duplicateInternal, duplicateSelf] = duplicate.shapes as ShapeModel[];

    expect(duplicateExternal!.hyperlink).toEqual(external.hyperlink);
    expect(duplicateInternal!.hyperlink).toEqual({
      slide: document.slides.indexOf(target) + 1,
    });
    expect(duplicateSelf!.hyperlink).toEqual({
      slide: document.slides.indexOf(duplicate) + 1,
      tooltip: '',
    });
    document.moveSlide(document.slides.indexOf(target), 0);
    expect(internal.hyperlink).toEqual({ slide: 1 });
    expect(duplicateInternal!.hyperlink).toEqual({ slide: 1 });
    expect(self.hyperlink).toEqual({
      slide: document.slides.indexOf(source) + 1,
      tooltip: '',
    });

    const beforeTarget = document.addSlide();
    document.moveSlide(document.slides.indexOf(beforeTarget), 0);
    expect(internal.hyperlink).toEqual({ slide: 2 });
    expect(duplicateInternal!.hyperlink).toEqual({ slide: 2 });
    document.deleteSlide(document.slides.indexOf(beforeTarget));
    expect(internal.hyperlink).toEqual({ slide: 1 });
    expect(duplicateInternal!.hyperlink).toEqual({ slide: 1 });

    document.deleteSlide(document.slides.indexOf(target));
    expect(internal.hyperlink).toBeUndefined();
    expect(duplicateInternal!.hyperlink).toBeUndefined();
    expect(external.hyperlink).toEqual({ url: 'https://example.com', tooltip: 'Visit' });
    expect(duplicateExternal!.hyperlink).toEqual(external.hyperlink);
    expect(self.hyperlink).toEqual({
      slide: document.slides.indexOf(source) + 1,
      tooltip: '',
    });
    expect(duplicateSelf!.hyperlink).toEqual({
      slide: document.slides.indexOf(duplicate) + 1,
      tooltip: '',
    });
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedSource = reopened.slides.find(({ partUri }) => partUri === source.partUri)!;
    const reopenedShapes = reopenedSource.shapes as ShapeModel[];
    expect(reopenedShapes.map(({ hyperlink }) => hyperlink)).toEqual([
      { url: 'https://example.com', tooltip: 'Visit' },
      undefined,
      { slide: reopened.slides.indexOf(reopenedSource) + 1, tooltip: '' },
    ]);
    await reopened.write();

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      const first = formatted.addSlide();
      formatted.addSlide();
      const shape = first.addShape('actionButtonForwardNext', {
        hyperlink: { slide: 2, tooltip: '' },
      });
      shape.hyperlink = { url: `https://example.com/${format}` };
      shape.hyperlink = { slide: 2, tooltip: '' };
      const formattedReopened = await PptxDocument.open(await formatted.write());
      expect(formattedReopened.format).toBe(format);
      expect((formattedReopened.slides[0]!.shapes[0] as ShapeModel).hyperlink)
        .toEqual({ slide: 2, tooltip: '' });
      await formattedReopened.write();
    }
  });

  it('creates preset shape fills through the public SDK surface', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const relationships = slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }));
    const runtimeUndefined: ShapeFill | undefined = undefined;
    const fill: {
      kind: 'solid';
      color: { kind: 'srgb'; value: string };
      transparency: number;
    } = {
      kind: 'solid',
      color: { kind: 'srgb', value: '#112233' },
      transparency: 50,
    };

    const omitted = slide.addShape('rect');
    const undefinedFill = slide.addShape('ellipse', { fill: runtimeUndefined } as never);
    const none = slide.addShape('star5', { fill: { kind: 'none' } });
    const solid = slide.addShape('diamond', { fill });
    const themed = slide.addShape('hexagon', {
      fill: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent3' },
        transparency: 25,
      },
    });
    fill.color.value = 'FFFFFF';
    fill.transparency = 0;

    expect(slide.shapes).toEqual([omitted, undefinedFill, none, solid, themed]);
    expect(slide.shapes[3]).toBe(solid);
    expect([omitted, undefinedFill, none, solid, themed].every(
      (shape) => shape instanceof ShapeModel,
    )).toBe(true);
    expect(slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }))).toEqual(relationships);

    const xml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect((xml.match(/<a:noFill\/>/g) ?? [])).toHaveLength(3);
    expect(xml).toContain(
      '<a:prstGeom prst="diamond"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:srgbClr val="112233"><a:alpha val="50000"/>' +
      '</a:srgbClr></a:solidFill><a:ln/>',
    );
    expect(xml).toContain(
      '<a:prstGeom prst="hexagon"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:schemeClr val="accent3"><a:alpha val="75000"/>' +
      '</a:schemeClr></a:solidFill><a:ln/>',
    );
  });

  it('creates preset shape lines through the public SDK surface', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const relationships = slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }));
    const runtimeUndefined: ShapeLine | undefined = undefined;
    const dash: ShapeLineDash = 'lgDashDotDot';
    const line: {
      kind: 'line';
      color: { kind: 'srgb'; value: string };
      transparency: number;
      width: number;
      dash: ShapeLineDash;
    } = {
      kind: 'line',
      color: { kind: 'srgb', value: '#112233' },
      transparency: 50,
      width: 2.5,
      dash,
    };

    const omitted = slide.addShape('rect');
    const undefinedLine = slide.addShape('ellipse', { line: runtimeUndefined } as never);
    const none = slide.addShape('star5', { line: { kind: 'none' } });
    const solid = slide.addShape('diamond', { line });
    const themed = slide.addShape('hexagon', {
      line: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent3' },
        transparency: 25,
        width: 0,
        dash: 'sysDash',
      },
    });
    line.color.value = 'FFFFFF';
    line.transparency = 0;
    line.width = 10;
    line.dash = 'solid';

    expect(slide.shapes).toEqual([omitted, undefinedLine, none, solid, themed]);
    expect(slide.shapes[3]).toBe(solid);
    expect([omitted, undefinedLine, none, solid, themed].every(
      (shape) => shape instanceof ShapeModel,
    )).toBe(true);
    expect(slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }))).toEqual(relationships);

    const xml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect((xml.match(/<a:ln\/>/g) ?? [])).toHaveLength(2);
    expect(xml).toContain(
      '<a:prstGeom prst="star5"><a:avLst/></a:prstGeom>' +
      '<a:noFill/><a:ln><a:noFill/></a:ln>',
    );
    expect(xml).toContain(
      '<a:prstGeom prst="diamond"><a:avLst/></a:prstGeom><a:noFill/>' +
      '<a:ln w="31750"><a:solidFill><a:srgbClr val="112233">' +
      '<a:alpha val="50000"/></a:srgbClr></a:solidFill>' +
      '<a:prstDash val="lgDashDotDot"/></a:ln>',
    );
    expect(xml).toContain(
      '<a:prstGeom prst="hexagon"><a:avLst/></a:prstGeom><a:noFill/>' +
      '<a:ln w="0"><a:solidFill><a:schemeClr val="accent3">' +
      '<a:alpha val="75000"/></a:schemeClr></a:solidFill>' +
      '<a:prstDash val="sysDash"/></a:ln>',
    );

    const invalidWidth: ShapeLine = {
      kind: 'line',
      color: { kind: 'srgb', value: 'FFFFFF' },
      // @ts-expect-error public shape line width is numeric points
      width: '2',
    };
    // @ts-expect-error public shape line dash union is closed
    const invalidDash: ShapeLineDash = 'dot';
    void [invalidWidth, invalidDash];
  });

  it('creates preset shape arrows through the public SDK surface', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const relationships = slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }));
    const begin: ShapeArrowType = 'triangle';
    const arrows: ShapeArrows = { begin, end: 'arrow' };
    const omitted = slide.addShape('line');
    const arrowOnly = slide.addShape('line', { arrows });
    const styled = slide.addShape('lineInv', {
      line: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent3' },
        width: 2.5,
        dash: 'sysDash',
      },
      arrows: { begin: 'none', end: 'stealth' },
    });

    expect(slide.shapes).toEqual([omitted, arrowOnly, styled]);
    expect(slide.shapes[1]).toBe(arrowOnly);
    expect([omitted, arrowOnly, styled].every((shape) => shape instanceof ShapeModel)).toBe(true);
    expect(slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }))).toEqual(relationships);

    const xml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(xml).toContain(
      '<a:ln><a:headEnd type="triangle"/><a:tailEnd type="arrow"/></a:ln>',
    );
    expect(xml).toContain(
      '<a:ln w="31750"><a:solidFill><a:schemeClr val="accent3"/>' +
      '</a:solidFill><a:prstDash val="sysDash"/>' +
      '<a:headEnd type="none"/><a:tailEnd type="stealth"/></a:ln>',
    );
    expect(xml).not.toContain('333333');

    // @ts-expect-error public shape arrow type union is closed
    const invalidType: ShapeArrowType = 'open';
    // @ts-expect-error public shape arrows expose begin/end, not PptxGenJS aliases
    const invalidAlias: ShapeArrows = { beginArrowType: 'arrow' };
    void [invalidType, invalidAlias];
  });

  it('creates preset shape shadows through the public SDK type and runtime surface', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const relationships = slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }));
    const runtimeUndefined: ShapeShadow | undefined = undefined;
    const color: { kind: 'srgb'; value: string } = { kind: 'srgb', value: '#123abc' };
    const shadow: ShapeShadow = {
      kind: 'outer',
      color,
      opacity: 0.42,
      blur: 7.25,
      angle: 123.4,
      distance: 5.5,
      rotateWithShape: true,
    };

    const omitted = slide.addShape('rect');
    const undefinedShadow = slide.addShape('ellipse', { shadow: runtimeUndefined } as never);
    const outer = slide.addShape('roundRect', { shadow });
    const inner = slide.addShape('star5', {
      shadow: {
        kind: 'inner',
        color: { kind: 'scheme', value: 'accent2' },
        opacity: 0,
        blur: 0,
        angle: 0,
        distance: 0,
      },
    });
    color.value = 'FFFFFF';

    expect(slide.shapes).toEqual([omitted, undefinedShadow, outer, inner]);
    expect(slide.shapes[2]).toBe(outer);
    expect(slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }))).toEqual(relationships);
    const xml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect((xml.match(/<a:effectLst>/g) ?? [])).toHaveLength(2);
    expect(xml).toContain(
      '<a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="bl" ' +
      'rotWithShape="1" blurRad="92075" dist="69850" dir="7404000">' +
      '<a:srgbClr val="123ABC"><a:alpha val="42000"/></a:srgbClr>' +
      '</a:outerShdw>',
    );
    expect(xml).toContain(
      '<a:innerShdw blurRad="0" dist="0" dir="0">' +
      '<a:schemeClr val="accent2"><a:alpha val="0"/></a:schemeClr>' +
      '</a:innerShdw>',
    );

    // @ts-expect-error public shape shadow kind union excludes none
    const invalidKind: ShapeShadow = { kind: 'none' };
    // @ts-expect-error inner shadows cannot rotate with the shape
    const invalidInnerRotate: ShapeShadow = { kind: 'inner', rotateWithShape: false };
    // @ts-expect-error public shape shadow uses kind, not PptxGenJS type
    const invalidAlias: ShapeShadow = { type: 'outer' };
    const invalidOpacity: ShapeShadow = {
      kind: 'outer',
      // @ts-expect-error public shape shadow opacity is numeric
      opacity: '0.5',
    };
    const invalidUnknown: ShapeShadow = {
      kind: 'outer',
      // @ts-expect-error public shape shadow union is closed
      offset: 4,
    };
    void [invalidKind, invalidInnerRotate, invalidAlias, invalidOpacity, invalidUnknown];
  });

  it('creates preset shape adjustments through the public SDK type and runtime surface', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const adjustments: ShapeAdjustment[] = [
      { name: 'adj1', value: 16_200_000 },
      { name: 'adj2', value: 0 },
      { name: 'adj3', value: 25_000 },
    ];
    const shape = slide.addShape('blockArc', { adjustments });
    adjustments[0] = { name: 'changed', value: 1 };

    expect(slide.shapes).toEqual([shape]);
    expect(slide.shapes[0]).toBe(shape);
    const xml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(xml).toContain(
      '<a:prstGeom prst="blockArc"><a:avLst>' +
      '<a:gd name="adj1" fmla="val 16200000"/>' +
      '<a:gd name="adj2" fmla="val 0"/>' +
      '<a:gd name="adj3" fmla="val 25000"/></a:avLst></a:prstGeom>',
    );

    // @ts-expect-error adjustment value must be numeric
    const invalidValue: ShapeAdjustment = { name: 'adj', value: '1' };
    // @ts-expect-error adjustment requires a name
    const missingName: ShapeAdjustment = { value: 1 };
    // @ts-expect-error adjustment values do not expose raw formulas
    const invalidFormula: ShapeAdjustment = { name: 'adj', value: 1, formula: 'val 1' };
    void [invalidValue, missingName, invalidFormula];
  });

  it('edits preset shape adjustments across duplicate, rollback, type reset, reopen, and all formats', async () => {
    const document = PptxDocument.create();
    const source = document.addSlide();
    const shape = source.addShape('blockArc', {
      adjustments: [
        { name: 'adj1', value: 16_200_000 },
        { name: 'adj2', value: 0 },
        { name: 'adj3', value: 25_000 },
      ],
    });
    expect(shape.adjustments).toEqual([
      { name: 'adj1', value: 16_200_000 },
      { name: 'adj2', value: 0 },
      { name: 'adj3', value: 25_000 },
    ]);
    expect(Object.isFrozen(shape.adjustments)).toBe(true);
    expect(shape.adjustments?.every(Object.isFrozen)).toBe(true);

    const duplicate = document.duplicateSlide(0);
    const duplicateShape = duplicate.shapes[0] as ShapeModel;
    duplicateShape.adjustments = [
      { name: 'adj1', value: 10_800_000 },
      { name: 'adj2', value: 5_400_000 },
    ];
    expect(shape.adjustments).toEqual([
      { name: 'adj1', value: 16_200_000 },
      { name: 'adj2', value: 0 },
      { name: 'adj3', value: 25_000 },
    ]);
    expect(duplicateShape.adjustments).toEqual([
      { name: 'adj1', value: 10_800_000 },
      { name: 'adj2', value: 5_400_000 },
    ]);

    expect(() => document.transaction(() => {
      shape.adjustments = [{ name: 'adj', value: 7 }];
      duplicateShape.adjustments = [];
      throw new Error('restore adjustment lifecycle');
    })).toThrow('restore adjustment lifecycle');
    expect(shape.adjustments?.[0]?.value).toBe(16_200_000);
    expect(duplicateShape.adjustments?.[0]?.value).toBe(10_800_000);
    expect(source.shapes[0]).toBe(shape);
    expect(duplicate.shapes[0]).toBe(duplicateShape);

    shape.presetType = 'blockArc';
    expect(shape.adjustments?.[2]?.value).toBe(25_000);
    shape.presetType = 'pie';
    expect(shape.adjustments).toEqual([]);

    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.slides[0]!.shapes[0] as ShapeModel).adjustments).toEqual([]);
    expect((reopened.slides[1]!.shapes[0] as ShapeModel).adjustments).toEqual([
      { name: 'adj1', value: 10_800_000 },
      { name: 'adj2', value: 5_400_000 },
    ]);

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      const formattedShape = formatted.addSlide().addShape('pie', {
        adjustments: [
          { name: 'adj1', value: 16_200_000 },
          { name: 'adj2', value: 0 },
        ],
      });
      formattedShape.adjustments = [
        { name: 'adj1', value: 10_800_000 },
        { name: 'adj2', value: 5_400_000 },
      ];
      const formattedReopened = await PptxDocument.open(await formatted.write());
      expect(formattedReopened.format).toBe(format);
      expect((formattedReopened.slides[0]!.shapes[0] as ShapeModel).adjustments).toEqual([
        { name: 'adj1', value: 10_800_000 },
        { name: 'adj2', value: 5_400_000 },
      ]);
    }
  });

  it('supports the shape shadow lifecycle across duplicate, move, rollback, reopen, and all formats', async () => {
    const document = PptxDocument.create();
    const source = document.addSlide();
    const shape = source.addShape('rect', {
      shadow: {
        kind: 'outer',
        color: { kind: 'srgb', value: 'AA0000' },
        opacity: 0.4,
        blur: 6,
        angle: 90,
        distance: 3,
        rotateWithShape: true,
      },
    });
    const initial = shape.shadow;
    expect(initial).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: 'AA0000' },
      opacity: 0.4,
      blur: 6,
      angle: 90,
      distance: 3,
      rotateWithShape: true,
    });
    expect(Object.isFrozen(initial)).toBe(true);
    expect(Object.isFrozen(initial?.color)).toBe(true);

    const sourcePart = document.opcPackage.requirePart(source.partUri);
    document.opcPackage.setPart(
      source.partUri,
      new TextDecoder().decode(sourcePart.bytes).replace(
        '<a:effectLst><a:outerShdw',
        '<a:effectLst><a:glow rad="12700"><a:srgbClr val="00FF00"/>' +
        '</a:glow><a:outerShdw',
      ),
      sourcePart.contentType,
    );

    const duplicate = document.duplicateSlide(0);
    const duplicateShape = duplicate.shapes[0] as ShapeModel;
    duplicateShape.shadow = {
      kind: 'inner',
      color: { kind: 'scheme', value: 'accent2' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
    };
    document.moveSlide(document.slides.indexOf(duplicate), 0);
    expect(shape.shadow).toEqual(initial);
    expect(duplicateShape.shadow?.kind).toBe('inner');
    for (const slide of [source, duplicate]) {
      expect(new TextDecoder().decode(
        document.opcPackage.requirePart(slide.partUri).bytes,
      )).toContain(
        '<a:glow rad="12700"><a:srgbClr val="00FF00"/></a:glow>',
      );
    }

    expect(() => document.transaction(() => {
      shape.shadow = { kind: 'outer', color: { kind: 'scheme', value: 'accent6' } };
      duplicateShape.shadow = undefined;
      throw new Error('restore shadow lifecycle');
    })).toThrow('restore shadow lifecycle');
    expect(shape.shadow).toEqual(initial);
    expect(duplicateShape.shadow?.kind).toBe('inner');
    expect(source.shapes[0]).toBe(shape);
    expect(duplicate.shapes[0]).toBe(duplicateShape);

    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.slides[0]!.shapes[0] as ShapeModel).shadow).toEqual({
      kind: 'inner',
      color: { kind: 'scheme', value: 'accent2' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
    });
    expect((reopened.slides[1]!.shapes[0] as ShapeModel).shadow).toEqual(initial);
    const secondWrite = await PptxDocument.open(await reopened.write());
    expect((secondWrite.slides[1]!.shapes[0] as ShapeModel).shadow).toEqual(initial);

    document.deleteSlide(document.slides.indexOf(source));
    expect(() => shape.shadow).toThrow();

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      const formattedShape = formatted.addSlide().addShape('hexagon', {
        shadow: { kind: 'outer' },
      });
      formattedShape.shadow = {
        kind: 'inner',
        color: { kind: 'scheme', value: 'accent4' },
        opacity: 0.5,
        blur: 2,
        angle: 45,
        distance: 1,
      };
      const formattedReopened = await PptxDocument.open(await formatted.write());
      expect(formattedReopened.format).toBe(format);
      expect((formattedReopened.slides[0]!.shapes[0] as ShapeModel).shadow).toEqual({
        kind: 'inner',
        color: { kind: 'scheme', value: 'accent4' },
        opacity: 0.5,
        blur: 2,
        angle: 45,
        distance: 1,
      });
    }
  });

  it('preserves editable shape arrows across duplicate, rollback, reopen, and all formats', async () => {
    const document = PptxDocument.create();
    const source = document.addSlide();
    const relationships = source.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }));
    const shape = source.addShape('line', {
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: 'AA0000' },
        width: 2.5,
        dash: 'lgDashDot',
      },
      arrows: { begin: 'triangle', end: 'arrow' },
    });
    const text = source.addText('Shape arrow text');
    text.arrows = { end: 'stealth' };
    expect(shape.arrows).toEqual({ begin: 'triangle', end: 'arrow' });
    expect(text.arrows).toEqual({ end: 'stealth' });

    const duplicate = document.duplicateSlide(0);
    const duplicateShape = duplicate.shapes[0] as ShapeModel;
    const duplicateText = duplicate.shapes[1] as ShapeModel;
    duplicateShape.arrows = { begin: 'diamond' };
    duplicateShape.line = undefined;
    duplicateText.arrows = undefined;
    expect(shape.arrows).toEqual({ begin: 'triangle', end: 'arrow' });
    expect(text.arrows).toEqual({ end: 'stealth' });
    expect(duplicateShape.arrows).toEqual({ begin: 'diamond' });
    expect(duplicateShape.line).toBeUndefined();
    expect(duplicateText.arrows).toBeUndefined();

    expect(() => document.transaction(() => {
      shape.arrows = { begin: 'none', end: 'oval' };
      expect(shape.arrows).toEqual({ begin: 'none', end: 'oval' });
      throw new Error('restore shape arrows');
    })).toThrow('restore shape arrows');
    expect(shape.arrows).toEqual({ begin: 'triangle', end: 'arrow' });
    expect(source.shapes[0]).toBe(shape);
    expect(source.shapes[1]).toBe(text);
    expect(source.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }))).toEqual(relationships);

    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.slides[0]!.shapes[0] as ShapeModel).arrows)
      .toEqual({ begin: 'triangle', end: 'arrow' });
    expect((reopened.slides[0]!.shapes[1] as ShapeModel).arrows)
      .toEqual({ end: 'stealth' });
    expect((reopened.slides[1]!.shapes[0] as ShapeModel).arrows)
      .toEqual({ begin: 'diamond' });
    expect((reopened.slides[1]!.shapes[1] as ShapeModel).arrows).toBeUndefined();

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      const formattedSlide = formatted.addSlide();
      const formattedShape = formattedSlide.addShape('line', {
        line: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent4' },
          width: 3,
          dash: 'lgDashDotDot',
        },
        arrows: { begin: 'none', end: 'triangle' },
      });
      const formattedText = formattedSlide.addText('Editable existing text arrows');
      formattedText.arrows = { begin: 'oval', end: 'stealth' };
      const formattedReopened = await PptxDocument.open(await formatted.write());
      expect(formattedReopened.format).toBe(format);
      expect((formattedReopened.slides[0]!.shapes[0] as ShapeModel).arrows)
        .toEqual({ begin: 'none', end: 'triangle' });
      expect((formattedReopened.slides[0]!.shapes[1] as ShapeModel).arrows)
        .toEqual({ begin: 'oval', end: 'stealth' });
      expect(formattedShape.line).toEqual({
        kind: 'line',
        color: { kind: 'scheme', value: 'accent4' },
        width: 3,
        dash: 'lgDashDotDot',
      });
    }
  });

  it('preserves editable shape lines across duplicate, rollback, reopen, and all formats', async () => {
    const document = PptxDocument.create();
    const source = document.addSlide();
    const relationships = source.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }));
    const shape = source.addShape('rect', {
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: 'AA0000' },
        transparency: 10,
        width: 2.5,
        dash: 'lgDashDot',
      },
    });
    const text = source.addText('Shape line text');
    expect(shape.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: 'AA0000' },
      transparency: 10,
      width: 2.5,
      dash: 'lgDashDot',
    });
    expect(text.line).toEqual({ kind: 'none' });

    text.line = {
      kind: 'line',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 35,
    };
    const duplicate = document.duplicateSlide(0);
    const duplicateShape = duplicate.shapes[0] as ShapeModel;
    const duplicateText = duplicate.shapes[1] as ShapeModel;
    duplicateShape.line = { kind: 'none' };
    duplicateText.line = undefined;
    expect(shape.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: 'AA0000' },
      transparency: 10,
      width: 2.5,
      dash: 'lgDashDot',
    });
    expect(text.line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 35,
      width: 1,
      dash: 'solid',
    });
    expect(duplicateShape.line).toEqual({ kind: 'none' });
    expect(duplicateText.line).toBeUndefined();

    expect(() => document.transaction(() => {
      shape.line = {
        kind: 'line',
        color: { kind: 'srgb', value: '00FF00' },
        width: 0,
        dash: 'sysDash',
      };
      expect(shape.line).toEqual({
        kind: 'line',
        color: { kind: 'srgb', value: '00FF00' },
        width: 0,
        dash: 'sysDash',
      });
      throw new Error('restore shape line');
    })).toThrow('restore shape line');
    expect(shape.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: 'AA0000' },
      transparency: 10,
      width: 2.5,
      dash: 'lgDashDot',
    });
    expect(source.shapes[0]).toBe(shape);
    expect(source.shapes[1]).toBe(text);
    expect(source.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }))).toEqual(relationships);

    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.slides[0]!.shapes[0] as ShapeModel).line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: 'AA0000' },
      transparency: 10,
      width: 2.5,
      dash: 'lgDashDot',
    });
    expect((reopened.slides[0]!.shapes[1] as ShapeModel).line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 35,
      width: 1,
      dash: 'solid',
    });
    expect((reopened.slides[1]!.shapes[0] as ShapeModel).line)
      .toEqual({ kind: 'none' });
    expect((reopened.slides[1]!.shapes[1] as ShapeModel).line).toBeUndefined();

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      const formattedSlide = formatted.addSlide();
      formattedSlide.addShape('hexagon', {
        line: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent4' },
          transparency: 50,
          width: 3,
          dash: 'lgDashDotDot',
        },
      });
      const formattedText = formattedSlide.addText('Editable existing text line');
      formattedText.line = { kind: 'none' };
      const formattedReopened = await PptxDocument.open(await formatted.write());
      expect(formattedReopened.format).toBe(format);
      expect((formattedReopened.slides[0]!.shapes[0] as ShapeModel).line).toEqual({
        kind: 'line',
        color: { kind: 'scheme', value: 'accent4' },
        transparency: 50,
        width: 3,
        dash: 'lgDashDotDot',
      });
      expect((formattedReopened.slides[0]!.shapes[1] as ShapeModel).line)
        .toEqual({ kind: 'none' });
    }
  });

  it('preserves editable shape fills across duplicate, rollback, reopen, and all formats', async () => {
    const document = PptxDocument.create();
    const source = document.addSlide();
    const relationships = source.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }));
    const shape = source.addShape('rect', {
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'AA0000' },
        transparency: 10,
      },
    });
    const text = source.addText('Shape fill text');
    expect(shape.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'AA0000' },
      transparency: 10,
    });
    expect(text.fill).toEqual({ kind: 'none' });

    text.fill = {
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 35,
    };
    const duplicate = document.duplicateSlide(0);
    const duplicateShape = duplicate.shapes[0] as ShapeModel;
    const duplicateText = duplicate.shapes[1] as ShapeModel;
    duplicateShape.fill = { kind: 'none' };
    duplicateText.fill = undefined;
    expect(shape.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'AA0000' },
      transparency: 10,
    });
    expect(text.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 35,
    });
    expect(duplicateShape.fill).toEqual({ kind: 'none' });
    expect(duplicateText.fill).toBeUndefined();

    expect(() => document.transaction(() => {
      shape.fill = {
        kind: 'solid',
        color: { kind: 'srgb', value: '00FF00' },
      };
      expect(shape.fill).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: '00FF00' },
      });
      throw new Error('restore shape fill');
    })).toThrow('restore shape fill');
    expect(shape.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'AA0000' },
      transparency: 10,
    });
    expect(source.shapes[0]).toBe(shape);
    expect(source.shapes[1]).toBe(text);
    expect(source.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }))).toEqual(relationships);

    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.slides[0]!.shapes[0] as ShapeModel).fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'AA0000' },
      transparency: 10,
    });
    expect((reopened.slides[0]!.shapes[1] as ShapeModel).fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 35,
    });
    expect((reopened.slides[1]!.shapes[0] as ShapeModel).fill).toEqual({ kind: 'none' });
    expect((reopened.slides[1]!.shapes[1] as ShapeModel).fill).toBeUndefined();

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      const formattedSlide = formatted.addSlide();
      formattedSlide.addShape('hexagon', {
        fill: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent4' },
          transparency: 50,
        },
      });
      const formattedText = formattedSlide.addText('Editable existing text');
      formattedText.fill = { kind: 'none' };
      const formattedReopened = await PptxDocument.open(await formatted.write());
      expect(formattedReopened.format).toBe(format);
      expect((formattedReopened.slides[0]!.shapes[0] as ShapeModel).fill).toEqual({
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent4' },
        transparency: 50,
      });
      expect((formattedReopened.slides[0]!.shapes[1] as ShapeModel).fill)
        .toEqual({ kind: 'none' });
    }
  });

  it('creates all 178 canonical preset shapes in catalog order', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const created = PRESET_SHAPE_TYPES.map((type) => slide.addShape(type));

    expect(created).toHaveLength(178);
    expect(created.map(({ id }) => id)).toEqual(
      Array.from({ length: 178 }, (_, index) => index + 2),
    );
    expect(created.map(({ name }) => name)).toEqual(
      Array.from({ length: 178 }, (_, index) => `Shape ${index + 2}`),
    );
    expect(created.map(({ presetType }) => presetType)).toEqual(PRESET_SHAPE_TYPES);
    expect(new Set(created)).toHaveLength(178);
    expect(slide.shapes).toEqual(created);

    const xml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    const tokens = [...xml.matchAll(/<a:prstGeom prst="([^"]+)"/g)]
      .map((match) => match[1]);
    expect(tokens).toEqual(PRESET_SHAPE_TYPES);
  }, 15_000);

  it('inserts preset shapes before extLst and rejects invalid additions without mutation', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const part = document.opcPackage.requirePart(slide.partUri);
    const withExtension = new TextDecoder().decode(part.bytes).replace(
      '</p:spTree>',
      '<p:extLst><p:ext uri="urn:test"><x:opaque xmlns:x="urn:test">KEEP</x:opaque>' +
      '</p:ext></p:extLst></p:spTree>',
    );
    document.opcPackage.setPart(slide.partUri, withExtension, part.contentType);

    slide.addShape('flowChartDecision');
    slide.addShape('actionButtonHome');
    const updated = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(updated.indexOf('prst="actionButtonHome"')).toBeLessThan(updated.indexOf('<p:extLst>'));
    expect(updated).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');

    const before = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const journal = [...document.opcPackage.mutations];
    for (const operation of [
      () => slide.addShape('folderCorner' as never),
      () => slide.addShape('custGeom' as never),
      () => slide.addShape('rect', { width: 0 as never }),
      () => slide.addShape('rect', { rotation: degrees(361) as never }),
      () => slide.addShape('rect', { flipHorizontal: 1 as never }),
      () => slide.addShape('rect', { unknown: true } as never),
    ]) expect(operation).toThrow();
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
  });

  it('preserves preset shapes through duplicate, move, rollback, reopen, and all formats', async () => {
    const document = PptxDocument.create();
    document.addSection({ title: 'Shapes' });
    const source = document.addSlide({ sectionTitle: 'Shapes' });
    source.hidden = true;
    const rectangle = source.addShape('rect', { name: 'Source rectangle' });
    source.addShape('foldedCorner', { x: inches(2), rotation: degrees(-30) });
    const duplicate = document.duplicateSlide(0);
    document.moveSlide(1, 0);

    let rolledBack: ShapeModel | undefined;
    expect(() => document.transaction(() => {
      rolledBack = source.addShape('star5');
      throw new Error('restore preset shape');
    })).toThrow('restore preset shape');
    expect(source.shapes).toHaveLength(2);
    expect(source.shapes[0]).toBe(rectangle);
    expect(() => rolledBack!.name).toThrow(ModelParseError);

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides.map(({ hidden }) => hidden)).toEqual([true, true]);
    const reopenedSection = reopened.sections?.[0];
    expect(reopenedSection).toBeDefined();
    expect(reopenedSection!.slideIds).toEqual(
      reopened.slides.map(({ slideId }) => slideId),
    );
    for (const slide of reopened.slides) {
      const xml = new TextDecoder().decode(reopened.opcPackage.requirePart(slide.partUri).bytes);
      expect([...xml.matchAll(/<a:prstGeom prst="([^"]+)"/g)].map((match) => match[1]))
        .toEqual(['rect', 'foldedCorner']);
      expect(slide.shapes.map(({ name }) => name)).toEqual(['Source rectangle', 'Shape 3']);
    }
    expect(duplicate.partUri).not.toBe(source.partUri);

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      formatted.addSlide().addShape('hexagon', { rotation: degrees(15) });
      const formattedReopened = await PptxDocument.open(await formatted.write());
      expect(formattedReopened.format).toBe(format);
      const shape = formattedReopened.slides[0]?.shapes[0];
      expect(shape?.name).toBe('Shape 2');
      expect(shape?.transform.rotation).toBe(degrees(15));
    }
  });

  it('reads and replaces preset types without changing unrelated shape content or identity', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addShape('rect', { name: 'Editable geometry' });
    const text = slide.addText('Keep text', { name: 'Text geometry' });
    const part = document.opcPackage.requirePart(slide.partUri);
    const customized = new TextDecoder().decode(part.bytes)
      .replace(
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/>',
        '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 7"/></a:avLst>' +
        '<x:old xmlns:x="urn:test"/></a:prstGeom>' +
        '<a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill><a:ln w="9"/>',
      );
    document.opcPackage.setPart(slide.partUri, customized, part.contentType);

    expect(shape.presetType).toBe('rect');
    expect(text.presetType).toBe('rect');
    const beforeNoOp = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...document.opcPackage.mutations];
    shape.presetType = 'rect';
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeNoOp);
    expect(document.opcPackage.mutations).toEqual(noOpJournal);

    shape.presetType = 'ellipse';
    text.presetType = 'roundRect';
    expect(shape.presetType).toBe('ellipse');
    expect(text.presetType).toBe('roundRect');
    expect(text.text).toBe('Keep text');
    expect(slide.shapes[0]).toBe(shape);
    expect(slide.shapes[1]).toBe(text);
    const updated = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(updated).toContain('<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>');
    expect(updated).toContain('<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>');
    expect(updated).not.toContain('name="adj"');
    expect(updated).not.toContain('<x:old');
    expect(updated).toContain('<a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>');
    expect(updated).toContain('<a:ln w="9"/>');
    expect(updated).toContain('<a:t xml:space="preserve">Keep text</a:t>');
  });

  it('isolates preset type edits across duplicates, rollback, malformed input, and reopen', async () => {
    const document = PptxDocument.create();
    const source = document.addSlide();
    const sourceShape = source.addShape('rect');
    const duplicate = document.duplicateSlide(0);
    const duplicateShape = duplicate.shapes[0] as ShapeModel;
    duplicateShape.presetType = 'star5';
    expect(sourceShape.presetType).toBe('rect');
    expect(duplicateShape.presetType).toBe('star5');

    expect(() => document.transaction(() => {
      sourceShape.presetType = 'hexagon';
      expect(sourceShape.presetType).toBe('hexagon');
      throw new Error('restore preset type');
    })).toThrow('restore preset type');
    expect(sourceShape.presetType).toBe('rect');
    expect(source.shapes[0]).toBe(sourceShape);

    const beforeInvalid = document.opcPackage.requirePart(source.partUri).bytes.slice();
    const invalidJournal = [...document.opcPackage.mutations];
    for (const value of ['folderCorner', 'custGeom', 'unknown', 7, null]) {
      expect(() => {
        sourceShape.presetType = value as never;
      }).toThrow(TypeError);
    }
    expect(document.opcPackage.requirePart(source.partUri).bytes).toEqual(beforeInvalid);
    expect(document.opcPackage.mutations).toEqual(invalidJournal);

    const sourcePart = document.opcPackage.requirePart(source.partUri);
    const malformed = new TextDecoder().decode(sourcePart.bytes)
      .replace('prst="rect"', 'prst="folderCorner"');
    document.opcPackage.setPart(source.partUri, malformed, sourcePart.contentType);
    expect(sourceShape.presetType).toBeUndefined();
    const beforeMalformedEdit = document.opcPackage.requirePart(source.partUri).bytes.slice();
    const malformedJournal = [...document.opcPackage.mutations];
    expect(() => {
      sourceShape.presetType = 'ellipse';
    }).toThrow(ModelParseError);
    expect(document.opcPackage.requirePart(source.partUri).bytes).toEqual(beforeMalformedEdit);
    expect(document.opcPackage.mutations).toEqual(malformedJournal);
    sourceShape.setTransform({ x: inches(3) });
    expect(new TextDecoder().decode(document.opcPackage.requirePart(source.partUri).bytes))
      .toContain('prst="folderCorner"');

    const reopened = await PptxDocument.open(await document.write());
    const reopenedSourceShape = reopened.slides[0]?.shapes[0];
    const reopenedDuplicateShape = reopened.slides[1]?.shapes[0];
    expect(reopenedSourceShape).toBeInstanceOf(ShapeModel);
    expect(reopenedDuplicateShape).toBeInstanceOf(ShapeModel);
    expect((reopenedSourceShape as ShapeModel).presetType).toBeUndefined();
    expect((reopenedDuplicateShape as ShapeModel).presetType).toBe('star5');
  });

  it('preserves hidden slide state through lifecycle, rollback, and all formats', async () => {
    const document = PptxDocument.create();
    const visible = document.addSlide();
    const hidden = document.addSlide();
    hidden.hidden = true;
    const section = document.addSection({ title: 'Hidden section' });
    document.assignSlideToSection(1, section.id);

    const duplicate = document.duplicateSlide(1);
    expect(duplicate.hidden).toBe(true);
    expect(document.sections).toEqual([
      { ...section, slideIds: [hidden.slideId, duplicate.slideId] },
    ]);
    document.moveSlide(document.slides.indexOf(duplicate), 0);
    expect(document.slides.map(({ hidden: state }) => state)).toEqual([true, false, true]);
    expect(document.sections).toEqual([
      { ...section, slideIds: [duplicate.slideId, hidden.slideId] },
    ]);
    document.deleteSlide(document.slides.indexOf(hidden));
    expect(document.slides.map(({ hidden: state }) => state)).toEqual([true, false]);
    expect(document.slides[0]).toBe(duplicate);
    expect(document.slides[1]).toBe(visible);
    expect(document.sections).toEqual([
      { ...section, slideIds: [duplicate.slideId] },
    ]);

    const beforeRollback = new Map(
      document.opcPackage.parts.map(({ uri, bytes }) => [uri, bytes.slice()]),
    );
    const rollbackJournal = [...document.opcPackage.mutations];
    const rollbackSlides = [...document.slides];
    const rollbackSections = document.sections;
    expect(() => document.transaction((draft) => {
      draft.slides[0]!.hidden = false;
      draft.slides[1]!.hidden = true;
      const temporary = draft.duplicateSlide(0);
      draft.moveSlide(draft.slides.indexOf(temporary), 0);
      draft.deleteSlide(1);
      throw new Error('rollback hidden slide lifecycle');
    })).toThrow('rollback hidden slide lifecycle');
    expect(document.opcPackage.parts.map(({ uri }) => uri)).toEqual([...beforeRollback.keys()]);
    for (const { uri, bytes } of document.opcPackage.parts) {
      expect(bytes).toEqual(beforeRollback.get(uri));
    }
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(document.slides).toEqual(rollbackSlides);
    expect(document.slides[0]).toBe(rollbackSlides[0]);
    expect(document.slides[1]).toBe(rollbackSlides[1]);
    expect(document.slides.map(({ hidden: state }) => state)).toEqual([true, false]);
    expect(document.sections).toEqual(rollbackSections);

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides.map(({ hidden: state }) => state)).toEqual([true, false]);
    expect(reopened.sections).toEqual(document.sections);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      formatted.addSlide();
      formatted.addSlide().hidden = true;
      const formattedReopened = await PptxDocument.open(await formatted.write());
      expect(formattedReopened.format).toBe(format);
      expect(formattedReopened.slides.map(({ hidden: state }) => state)).toEqual([false, true]);
      expect(validatePackage(formattedReopened.opcPackage)
        .filter(({ severity }) => severity === 'error')).toEqual([]);
    }
  }, 20_000);

  it('preserves speaker notes through lifecycle, rollback, sections, hidden state, and all formats', async () => {
    const document = PptxDocument.create();
    const visible = document.addSlide().addNotes('Visible notes');
    const hidden = document.addSlide().addNotes('Hidden notes');
    hidden.hidden = true;
    const empty = document.addSlide();
    empty.notes = '';
    const lazy = document.addSlide();
    const section = document.addSection({ title: 'Noted section' });
    document.assignSlideToSection(document.slides.indexOf(hidden), section.id);

    const sourceNotesRelationship = hidden.relationships.find(
      ({ type }) => type.endsWith('/notesSlide'),
    )!;
    const sourceNotesUri = sourceNotesRelationship.resolvedTarget!;
    const sourceMasterUri = document.opcPackage.relationships(sourceNotesUri).find(
      ({ type }) => type.endsWith('/notesMaster'),
    )!.resolvedTarget!;

    const duplicate = document.duplicateSlide(document.slides.indexOf(hidden));
    expect(duplicate.notes).toBe('Hidden notes');
    expect(duplicate.hidden).toBe(true);
    const duplicateNotesUri = duplicate.relationships.find(
      ({ type }) => type.endsWith('/notesSlide'),
    )!.resolvedTarget!;
    expect(duplicateNotesUri).not.toBe(sourceNotesUri);
    expect(document.opcPackage.relationships(duplicateNotesUri).find(
      ({ type }) => type.endsWith('/slide'),
    )?.resolvedTarget).toBe(duplicate.partUri);
    expect(document.opcPackage.relationships(duplicateNotesUri).find(
      ({ type }) => type.endsWith('/notesMaster'),
    )?.resolvedTarget).toBe(sourceMasterUri);
    expect(document.sections).toEqual([
      { ...section, slideIds: [hidden.slideId, duplicate.slideId] },
    ]);

    document.moveSlide(document.slides.indexOf(duplicate), 0);
    duplicate.notes = 'Duplicate notes';
    document.deleteSlide(document.slides.indexOf(hidden));
    expect(document.slides).toEqual([duplicate, visible, empty, lazy]);
    expect(document.slides.map(({ notes }) => notes)).toEqual([
      'Duplicate notes',
      'Visible notes',
      '',
      undefined,
    ]);
    expect(document.slides.map(({ hidden: state }) => state)).toEqual([true, false, false, false]);
    expect(document.sections).toEqual([
      { ...section, slideIds: [duplicate.slideId] },
    ]);
    expect(document.opcPackage.hasPart(sourceNotesUri)).toBe(false);
    expect(document.opcPackage.hasPart(duplicateNotesUri)).toBe(true);
    expect(document.opcPackage.hasPart(sourceMasterUri)).toBe(true);

    const beforeRollback = new Map(
      document.opcPackage.parts.map(({ uri, bytes }) => [uri, bytes.slice()]),
    );
    const rollbackJournal = [...document.opcPackage.mutations];
    const rollbackSlides = [...document.slides];
    const rollbackSections = document.sections;
    expect(() => document.transaction((draft) => {
      draft.slides[0]!.notes = 'Temporary edit';
      draft.slides[1]!.notes = undefined;
      draft.slides[2]!.addNotes('Temporary empty edit');
      const temporary = draft.duplicateSlide(0);
      temporary.notes = 'Temporary duplicate';
      draft.moveSlide(draft.slides.indexOf(temporary), 0);
      draft.deleteSlide(2);
      throw new Error('rollback speaker notes lifecycle');
    })).toThrow('rollback speaker notes lifecycle');
    expect(document.opcPackage.parts.map(({ uri }) => uri)).toEqual([...beforeRollback.keys()]);
    for (const { uri, bytes } of document.opcPackage.parts) {
      expect(bytes).toEqual(beforeRollback.get(uri));
    }
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(document.slides).toEqual(rollbackSlides);
    expect(document.slides[0]).toBe(rollbackSlides[0]);
    expect(document.slides[1]).toBe(rollbackSlides[1]);
    expect(document.slides.map(({ notes }) => notes)).toEqual([
      'Duplicate notes',
      'Visible notes',
      '',
      undefined,
    ]);
    expect(document.sections).toEqual(rollbackSections);

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides.map(({ notes }) => notes)).toEqual([
      'Duplicate notes',
      'Visible notes',
      '',
      undefined,
    ]);
    expect(reopened.slides.map(({ hidden: state }) => state)).toEqual([true, false, false, false]);
    expect(reopened.sections).toEqual(document.sections);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const formatted = PptxDocument.create({ format });
      formatted.addSlide();
      formatted.addSlide().notes = '';
      formatted.addSlide().addNotes(`Notes ${format}`);
      const formattedReopened = await PptxDocument.open(await formatted.write());
      expect(formattedReopened.format).toBe(format);
      expect(formattedReopened.slides.map(({ notes }) => notes)).toEqual([
        undefined,
        '',
        `Notes ${format}`,
      ]);
      expect(validatePackage(formattedReopened.opcPackage)
        .filter(({ severity }) => severity === 'error')).toEqual([]);
    }
  }, 20_000);

  it('repairs a fully missing notes master and rolls the repair back atomically', () => {
    const removeNotesMaster = (document: PptxDocument): void => {
      const presentationPart = document.opcPackage.requirePart('/ppt/presentation.xml');
      document.opcPackage.setPart(
        presentationPart.uri,
        new TextDecoder().decode(presentationPart.bytes).replace(
          /<p:notesMasterIdLst>.*?<\/p:notesMasterIdLst>/,
          '',
        ),
        presentationPart.contentType,
      );
      document.opcPackage.deletePart('/ppt/notesMasters/notesMaster1.xml');
    };

    const document = PptxDocument.create();
    const slide = document.addSlide();
    removeNotesMaster(document);
    expect(document.opcPackage.parts.filter(
      ({ contentType }) => contentType.endsWith('.notesMaster+xml'),
    )).toHaveLength(0);
    slide.addNotes('Repaired master');
    expect(slide.notes).toBe('Repaired master');
    const masterRelationships = document.opcPackage.relationships('/ppt/presentation.xml').filter(
      ({ type }) => type.endsWith('/notesMaster'),
    );
    expect(masterRelationships).toHaveLength(1);
    const masterUri = masterRelationships[0]!.resolvedTarget!;
    expect(document.opcPackage.relationships(masterUri).find(
      ({ type }) => type.endsWith('/theme'),
    )?.resolvedTarget).toBe('/ppt/theme/theme1.xml');
    expect(new TextDecoder().decode(
      document.opcPackage.requirePart('/ppt/presentation.xml').bytes,
    )).toContain(`<p:notesMasterId r:id="${masterRelationships[0]!.id}"/>`);
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);

    const rollback = PptxDocument.create();
    const rollbackSlide = rollback.addSlide();
    removeNotesMaster(rollback);
    const before = rollback.opcPackage.parts.map(({ uri, bytes }) => ({ uri, bytes: bytes.slice() }));
    const journal = [...rollback.opcPackage.mutations];
    expect(() => rollback.transaction(() => {
      rollbackSlide.addNotes('Temporary repair');
      throw new Error('rollback notes master repair');
    })).toThrow('rollback notes master repair');
    expect(rollback.opcPackage.parts.map(({ uri, bytes }) => ({ uri, bytes }))).toEqual(before);
    expect(rollback.opcPackage.mutations).toEqual(journal);
    expect(rollbackSlide.notes).toBeUndefined();
  });

  it('creates, reads, and atomically edits detached presentation sections', async () => {
    const document = PptxDocument.create();
    const slides = [document.addSlide(), document.addSlide(), document.addSlide()];
    expect(document.sections).toEqual([]);

    const a = document.addSection({ title: 'A' });
    const c = document.addSection({ title: 'C' });
    const b = document.addSection({ title: 'B & <Two>', order: 1 });
    expect(document.sections?.map(({ title }) => title)).toEqual(['A', 'B & <Two>', 'C']);
    expect([a.id, b.id, c.id]).toEqual(expect.arrayContaining([
      expect.stringMatching(/^\{[0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$/),
      expect.stringMatching(/^\{[0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$/),
      expect.stringMatching(/^\{[0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$/),
    ]));
    expect(new Set([a.id, b.id, c.id]).size).toBe(3);

    const duplicateA = document.addSection({ title: 'A', order: 1 });
    document.renameSection(duplicateA.id, 'Second A');
    expect(document.sections?.map(({ title }) => title)).toEqual(['A', 'Second A', 'B & <Two>', 'C']);
    document.deleteSection(duplicateA.id);
    expect(document.sections?.map(({ title }) => title)).toEqual(['A', 'B & <Two>', 'C']);

    const detached = document.sections! as unknown as { title: string; slideIds: number[] }[];
    detached[0]!.title = 'Changed detached title';
    detached[0]!.slideIds.push(999);
    expect(document.sections?.[0]).toEqual({ id: a.id, title: 'A', slideIds: [] });

    document.assignSlideToSection(0, a.id);
    document.assignSlideToSection(1, b.id);
    document.assignSlideToSection(2, c.id);
    expect(document.sections?.map(({ slideIds }) => slideIds)).toEqual([
      [slides[0]!.slideId],
      [slides[1]!.slideId],
      [slides[2]!.slideId],
    ]);

    document.assignSlideToSection(1, a.id);
    document.assignSlideToSection(1, undefined);
    document.renameSection(b.id, 'B & <Renamed> "Q"');
    document.moveSection(c.id, 0);
    document.deleteSection(b.id);
    expect(document.sections).toEqual([
      { id: c.id, title: 'C', slideIds: [slides[2]!.slideId] },
      { id: a.id, title: 'A', slideIds: [slides[0]!.slideId] },
    ]);
    expect(document.slides).toEqual(slides);
    expect(new TextDecoder().decode(
      document.opcPackage.requirePart(document.presentationPartUri).bytes,
    )).not.toContain('name="B &amp; &lt;Renamed&gt; &quot;Q&quot;"');

    const beforeNoOp = document.opcPackage.requirePart(document.presentationPartUri).bytes;
    const journal = [...document.opcPackage.mutations];
    document.renameSection(c.id, 'C');
    document.moveSection(c.id, 0);
    document.assignSlideToSection(2, c.id);
    document.assignSlideToSection(1, undefined);
    expect(document.opcPackage.requirePart(document.presentationPartUri).bytes).toEqual(beforeNoOp);
    expect(document.opcPackage.mutations).toEqual(journal);

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.sections).toEqual(document.sections);
    expect(reopened.slides.map(({ slideId }) => slideId)).toEqual(slides.map(({ slideId }) => slideId));
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('rejects invalid and unsafe presentation section edits without mutation and rolls back', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const section = document.addSection({ title: 'Safe' });
    document.assignSlideToSection(0, section.id);
    const before = new Map(
      document.opcPackage.parts.map(({ uri, bytes }) => [uri, new Uint8Array(bytes)]),
    );
    const beforeJournal = [...document.opcPackage.mutations];
    const expectUnchanged = (): void => {
      expect(document.opcPackage.parts.map(({ uri }) => uri)).toEqual([...before.keys()]);
      for (const { uri, bytes } of document.opcPackage.parts) {
        expect(bytes).toEqual(before.get(uri));
      }
      expect(document.opcPackage.mutations).toEqual(beforeJournal);
    };
    const identity = document.slides[0];
    let accessorCalls = 0;
    const accessorOptions = Object.defineProperty({}, 'title', {
      enumerable: true,
      get: () => {
        accessorCalls += 1;
        return 'Accessor';
      },
    });
    const invalidOptions = [
      null,
      [],
      {},
      { title: '' },
      { title: '   ' },
      { title: 1 },
      { title: 'Safe', order: 1.5 },
      { title: 'Safe', order: 2 },
      { title: 'Safe', extra: true },
      { title: 'Safe', [Symbol('extra')]: true },
      Object.create({ title: 'Inherited' }),
      accessorOptions,
    ];
    for (const options of invalidOptions) {
      expect(() => document.addSection(options as never)).toThrow();
    }
    expect(accessorCalls).toBe(0);
    expect(() => document.renameSection('{00000000-0000-0000-0000-000000000099}', 'Missing')).toThrow(RangeError);
    expect(() => document.renameSection('bad', 'Bad')).toThrow(TypeError);
    expect(() => document.renameSection(section.id, '   ')).toThrow(TypeError);
    expect(() => document.moveSection(section.id, 1)).toThrow(RangeError);
    expect(() => document.deleteSection('{00000000-0000-0000-0000-000000000099}')).toThrow(RangeError);
    expect(() => document.assignSlideToSection(1, section.id)).toThrow(RangeError);
    expect(() => document.assignSlideToSection(0, '{00000000-0000-0000-0000-000000000099}')).toThrow(RangeError);
    expectUnchanged();
    expect(document.slides[0]).toBe(identity);
    expect(document.sections?.[0]?.slideIds).toEqual([slide.slideId]);

    expect(() => document.transaction((draft) => {
      const temporary = draft.addSection({ title: 'Temporary', order: 0 });
      draft.assignSlideToSection(0, temporary.id);
      draft.renameSection(section.id, 'Changed');
      throw new Error('rollback presentation sections');
    })).toThrow('rollback presentation sections');
    expectUnchanged();
    expect(document.slides[0]).toBe(identity);

    const part = document.opcPackage.requirePart(document.presentationPartUri);
    const source = new TextDecoder().decode(part.bytes);
    const member = `<p14:sldId id="${slide.slideId}"/>`;
    document.opcPackage.setPart(
      document.presentationPartUri,
      source.replace(member, `${member}${member}`),
      part.contentType,
    );
    const unsafe = document.opcPackage.requirePart(document.presentationPartUri).bytes;
    expect(document.sections).toBeUndefined();
    expect(() => document.renameSection(section.id, 'Unsafe')).toThrow(ModelParseError);
    expect(() => document.addSection({ title: 'Unsafe' })).toThrow(ModelParseError);
    expect(document.opcPackage.requirePart(document.presentationPartUri).bytes).toEqual(unsafe);
  });

  it('round-trips presentation section commands in all supported presentation formats', async () => {
    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const document = PptxDocument.create({ format });
      const slide = document.addSlide();
      const section = document.addSection({ title: `Section ${format}` });
      document.assignSlideToSection(0, section.id);
      const reopened = await PptxDocument.open(await document.write());
      expect(reopened.format).toBe(format);
      expect(reopened.sections).toEqual([
        { id: section.id, title: `Section ${format}`, slideIds: [slide.slideId] },
      ]);
    }
  });

  it('assigns added slides to explicit and automatic presentation sections', async () => {
    const document = PptxDocument.create();
    const loose = document.addSlide();
    expect(document.sections).toEqual([]);

    const firstIntro = document.addSection({ title: 'Intro' });
    const secondIntro = document.addSection({ title: 'Intro' });
    const escaped = document.addSection({ title: 'Data & <One>' });
    const explicitIntro = document.addSlide({ sectionTitle: 'Intro' });
    const explicitEscaped = document.addSlide({ sectionTitle: 'Data & <One>' });
    const automaticOne = document.addSlide();
    const automaticTwo = document.addSlide();
    const afterDefault = document.addSection({ title: 'After default' });
    const automaticThree = document.addSlide({ sectionTitle: undefined } as never);

    expect(document.sections).toEqual([
      { ...firstIntro, slideIds: [explicitIntro.slideId] },
      { ...secondIntro, slideIds: [] },
      { ...escaped, slideIds: [explicitEscaped.slideId] },
      expect.objectContaining({
        title: 'Default-1',
        slideIds: [automaticOne.slideId, automaticTwo.slideId],
      }),
      { ...afterDefault, slideIds: [] },
      expect.objectContaining({ title: 'Default-2', slideIds: [automaticThree.slideId] }),
    ]);
    expect(document.sections?.flatMap(({ slideIds }) => slideIds)).not.toContain(loose.slideId);

    let accessorCalls = 0;
    const accessorOptions = Object.defineProperty({}, 'sectionTitle', {
      enumerable: true,
      get: () => {
        accessorCalls += 1;
        return 'Intro';
      },
    });
    const beforeInvalid = new Map(
      document.opcPackage.parts.map(({ uri, bytes }) => [uri, bytes.slice()]),
    );
    const beforeInvalidJournal = [...document.opcPackage.mutations];
    const beforeInvalidSlides = [...document.slides];
    expect(() => document.addSlide({ sectionTitle: 'Missing' })).toThrow(RangeError);
    for (const options of [
      null,
      [],
      { sectionTitle: 1 },
      { sectionTitle: 'Intro', extra: true },
      { sectionTitle: 'Intro', [Symbol('extra')]: true },
      Object.create({ sectionTitle: 'Intro' }),
      accessorOptions,
    ]) {
      expect(() => document.addSlide(options as never)).toThrow(TypeError);
    }
    expect(accessorCalls).toBe(0);
    expect(document.opcPackage.parts.map(({ uri }) => uri)).toEqual([...beforeInvalid.keys()]);
    for (const { uri, bytes } of document.opcPackage.parts) {
      expect(bytes).toEqual(beforeInvalid.get(uri));
    }
    expect(document.opcPackage.mutations).toEqual(beforeInvalidJournal);
    expect(document.slides).toEqual(beforeInvalidSlides);

    const reopened = await PptxDocument.open(await document.write());
    const continued = reopened.addSlide();
    expect(reopened.sections?.at(-1)).toMatchObject({
      title: 'Default-2',
      slideIds: [automaticThree.slideId, continued.slideId],
    });
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('synchronizes duplicate, delete, and move slide presentation sections', async () => {
    const document = PptxDocument.create();
    const loose = document.addSlide();
    const aOne = document.addSlide();
    const bOne = document.addSlide();
    const aTwo = document.addSlide();
    const cOne = document.addSlide();
    aOne.addText('Assigned source');

    const a = document.addSection({ title: 'A' });
    const b = document.addSection({ title: 'B' });
    const c = document.addSection({ title: 'C' });
    document.assignSlideToSection(1, a.id);
    document.assignSlideToSection(2, b.id);
    document.assignSlideToSection(3, a.id);
    document.assignSlideToSection(4, c.id);

    const assignedDuplicate = document.duplicateSlide(1);
    const looseDuplicate = document.duplicateSlide(0);
    expect(assignedDuplicate.shapes[0]).toMatchObject({ text: 'Assigned source' });
    expect(document.sections).toEqual([
      { ...a, slideIds: [aOne.slideId, aTwo.slideId, assignedDuplicate.slideId] },
      { ...b, slideIds: [bOne.slideId] },
      { ...c, slideIds: [cOne.slideId] },
    ]);
    expect(document.sections?.flatMap(({ slideIds }) => slideIds)).not.toContain(loose.slideId);
    expect(document.sections?.flatMap(({ slideIds }) => slideIds)).not.toContain(looseDuplicate.slideId);

    document.deleteSlide(2);
    expect(document.sections?.find(({ id }) => id === b.id)).toEqual({ ...b, slideIds: [] });
    document.moveSlide(2, 0);
    document.moveSlide(document.slides.indexOf(assignedDuplicate), 1);
    expect(document.slides.map(({ slideId }) => slideId)).toEqual([
      aTwo.slideId,
      assignedDuplicate.slideId,
      loose.slideId,
      aOne.slideId,
      cOne.slideId,
      looseDuplicate.slideId,
    ]);
    expect(document.sections).toEqual([
      { ...a, slideIds: [aTwo.slideId, assignedDuplicate.slideId, aOne.slideId] },
      { ...b, slideIds: [] },
      { ...c, slideIds: [cOne.slideId] },
    ]);

    document.deleteSlide(document.slides.indexOf(aOne));
    expect(document.sections).toEqual([
      { ...a, slideIds: [aTwo.slideId, assignedDuplicate.slideId] },
      { ...b, slideIds: [] },
      { ...c, slideIds: [cOne.slideId] },
    ]);
    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.sections).toEqual(document.sections);
    expect(reopened.slides.map(({ slideId }) => slideId)).toEqual(
      document.slides.map(({ slideId }) => slideId),
    );
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('rejects unsafe section slide lifecycles and rolls composite operations back', () => {
    const unsafe = PptxDocument.create();
    const unsafeSlide = unsafe.addSlide();
    unsafe.addSlide();
    const unsafeSection = unsafe.addSection({ title: 'Unsafe' });
    unsafe.assignSlideToSection(0, unsafeSection.id);
    const part = unsafe.opcPackage.requirePart(unsafe.presentationPartUri);
    const member = `<p14:sldId id="${unsafeSlide.slideId}"/>`;
    unsafe.opcPackage.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(member, `${member}${member}`),
      part.contentType,
    );
    expect(unsafe.sections).toBeUndefined();
    const unsafeParts = new Map(
      unsafe.opcPackage.parts.map(({ uri, bytes }) => [uri, bytes.slice()]),
    );
    const unsafeJournal = [...unsafe.opcPackage.mutations];
    const unsafeSlides = [...unsafe.slides];
    for (const operation of [
      () => unsafe.addSlide({ sectionTitle: 'Unsafe' }),
      () => unsafe.duplicateSlide(0),
      () => unsafe.deleteSlide(0),
      () => unsafe.moveSlide(0, 1),
    ]) {
      expect(operation).toThrow(ModelParseError);
      expect(unsafe.opcPackage.parts.map(({ uri }) => uri)).toEqual([...unsafeParts.keys()]);
      for (const { uri, bytes } of unsafe.opcPackage.parts) {
        expect(bytes).toEqual(unsafeParts.get(uri));
      }
      expect(unsafe.opcPackage.mutations).toEqual(unsafeJournal);
      expect(unsafe.slides).toEqual(unsafeSlides);
    }

    const document = PptxDocument.create();
    const section = document.addSection({ title: 'Safe' });
    const first = document.addSlide({ sectionTitle: 'Safe' });
    const second = document.addSlide({ sectionTitle: 'Safe' });
    first.addText('Rollback source');
    const before = new Map(
      document.opcPackage.parts.map(({ uri, bytes }) => [uri, bytes.slice()]),
    );
    const journal = [...document.opcPackage.mutations];
    const identities = [...document.slides];
    expect(() => document.transaction((draft) => {
      draft.addSlide({ sectionTitle: 'Safe' });
      draft.duplicateSlide(0);
      draft.moveSlide(draft.slides.length - 1, 0);
      draft.deleteSlide(1);
      throw new Error('rollback section slide lifecycle');
    })).toThrow('rollback section slide lifecycle');
    expect(document.opcPackage.parts.map(({ uri }) => uri)).toEqual([...before.keys()]);
    for (const { uri, bytes } of document.opcPackage.parts) {
      expect(bytes).toEqual(before.get(uri));
    }
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(document.slides).toEqual(identities);
    expect(document.slides[0]).toBe(first);
    expect(document.slides[1]).toBe(second);
    expect(document.sections).toEqual([
      { ...section, slideIds: [first.slideId, second.slideId] },
    ]);
  });

  it('creates and reopens all native chart types through the public SDK in all six formats', async () => {
    const types = [
      'area', 'bar', 'bar3D', 'bubble', 'doughnut', 'line', 'pie', 'radar', 'scatter',
    ] as const;
    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const document = PptxDocument.create({ format });
      const slide = document.addSlide();
      const created: ChartModel[] = [];
      for (const [index, type] of types.entries()) {
        const series = type === 'scatter'
          ? [{ name: `${type} series`, xValues: [1, 2], values: [10, 20] }]
          : type === 'bubble'
            ? [{
                name: `${type} series`,
                xValues: [1, 2],
                values: [10, 20],
                sizes: [5, 6],
              }]
            : [{ name: `${type} series`, categories: ['Q1', 'Q2'], values: [10, 20] }];
        created.push(await document.addChart(0, type, series, index === 0 ? {
          name: `${format} chart`,
          altText: 'Quarterly chart',
          x: inches(2),
          y: inches(3),
          width: inches(7),
          height: inches(4),
        } : {}));
      }
      expect(slide.shapes.filter((shape) => shape instanceof ChartModel)).toEqual(created);
      expect(created[0]).toBeInstanceOf(ChartModel);
      expect(created[0]?.name).toBe(`${format} chart`);
      expect(created[0]?.altText).toBe('Quarterly chart');
      expect(created[0]?.transform).toMatchObject({
        x: inches(2),
        y: inches(3),
        width: inches(7),
        height: inches(4),
      });
      expect(created.every((chart) => chart.workbookPartUri !== undefined)).toBe(true);
      const combo = await document.addChart(0, [
        {
          type: 'bar',
          series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
        },
        {
          type: 'line',
          axis: 'secondary',
          series: [{ name: 'Trend', categories: ['Q1', 'Q2'], values: [11, 21] }],
        },
      ], { name: `${format} combo` });
      expect(combo.definition?.groups.map(({ type, axis }) => [type, axis])).toEqual([
        ['bar', 'primary'],
        ['line', 'secondary'],
      ]);
      await created[1]!.replaceSeries([{
        name: 'bar edited', categories: ['Q1', 'Q2'], values: [30, 40],
      }]);
      await created[3]!.replaceDefinition({ groups: [{
        type: 'bubble',
        series: [{
          name: 'bubble edited', xValues: [3, 4], values: [50, 60], sizes: [7, 8],
        }],
      }] });
      await created[8]!.replaceSeries([{
        name: 'scatter edited', xValues: [5, 6], values: [70, 80],
      }]);
      await combo.replaceDefinition({ groups: [
        {
          type: 'bar',
          series: [{ name: 'Revenue edited', categories: ['Q1', 'Q2'], values: [15, 25] }],
        },
        {
          type: 'line',
          axis: 'secondary',
          series: [{ name: 'Trend edited', categories: ['Q1', 'Q2'], values: [16, 26] }],
        },
      ] });
      const presentation = created[6]!;
      const presentationGroup = presentation.definition?.groups[0];
      if (presentationGroup?.type !== 'pie') throw new Error('Expected created pie chart');
      await presentation.replaceDefinition({
        groups: [{
          type: 'pie',
          series: presentationGroup.series,
          options: {
            firstSliceAngle: 45,
            dataLabels: {
              face: 'Aptos',
              size: 11,
              bold: true,
              color: { kind: 'scheme', value: 'accent1' },
              showCategoryName: true,
              showPercent: true,
              position: 'bestFit',
              numberFormat: '0%',
              showLeaderLines: true,
            },
          },
        }],
        options: {
          title: { text: `Presentation ${format}`, size: 18, bold: true },
          legend: { position: 'topRight', size: 10 },
          displayBlanksAs: 'span',
        },
      });
      const duplicate = document.duplicateSlide(0);
      const duplicateCharts = duplicate.shapes.filter(
        (shape): shape is ChartModel => shape instanceof ChartModel,
      );
      const duplicateCombo = duplicateCharts.at(-1)!;
      const duplicatePresentation = duplicateCharts[6]!;
      expect(duplicateCombo.definition).toEqual(combo.definition);
      expect(duplicateCombo.chartPartUri).not.toBe(combo.chartPartUri);
      expect(duplicateCombo.workbookPartUri).not.toBe(combo.workbookPartUri);
      expect(duplicatePresentation.definition).toEqual(presentation.definition);
      expect(duplicatePresentation.definition?.groups[0]?.options?.dataLabels).toMatchObject({
        showCategoryName: true,
        showPercent: true,
        position: 'bestFit',
      });
      const editedPresentation = presentation.definition!;
      const editedPresentationGroup = editedPresentation.groups[0];
      if (editedPresentationGroup?.type !== 'pie') throw new Error('Expected edited pie chart');
      await presentation.replaceDefinition({
        groups: [{
          type: 'pie',
          series: editedPresentationGroup.series,
          options: { firstSliceAngle: 45 },
        }],
        options: editedPresentation.options,
      });
      expect(presentation.definition?.groups[0]?.options?.dataLabels).toBeUndefined();
      expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);

      const reopened = await PptxDocument.open(await document.write());
      const charts = reopened.slides[0]!.shapes.filter(
        (shape): shape is ChartModel => shape instanceof ChartModel,
      );
      const singleCharts = charts.slice(0, types.length);
      const reopenedCombo = charts.at(-1)!;
      expect(reopened.format).toBe(format);
      expect(singleCharts[0]).toMatchObject({
        name: `${format} chart`,
        altText: 'Quarterly chart',
        transform: {
          x: inches(2),
          y: inches(3),
          width: inches(7),
          height: inches(4),
        },
      });
      expect(singleCharts.map((chart) => chart.definition?.groups[0]?.type)).toEqual(types);
      expect(singleCharts.map((chart) => chart.series[0]?.values)).toEqual(types.map((type) =>
        type === 'bar'
          ? [30, 40]
          : type === 'bubble'
            ? [50, 60]
            : type === 'scatter'
              ? [70, 80]
              : [10, 20]));
      expect(singleCharts.find((chart) => chart.definition?.groups[0]?.type === 'scatter')
        ?.series[0]?.xValues).toEqual([5, 6]);
      expect(singleCharts.find((chart) => chart.definition?.groups[0]?.type === 'bubble')
        ?.series[0]?.sizes).toEqual([7, 8]);
      expect(singleCharts[6]?.definition?.groups[0]?.options).toEqual({ firstSliceAngle: 45 });
      expect(singleCharts[6]?.definition?.options).toMatchObject({
        title: { text: `Presentation ${format}`, size: 18, bold: true },
        legend: { position: 'topRight', size: 10 },
        displayBlanksAs: 'span',
      });
      expect(reopenedCombo.definition?.groups.map(({ type, axis }) => [type, axis])).toEqual([
        ['bar', 'primary'],
        ['line', 'secondary'],
      ]);
      expect(reopenedCombo.series.map(({ name, values }) => ({ name, values }))).toEqual([
        { name: 'Revenue edited', values: [15, 25] },
        { name: 'Trend edited', values: [16, 26] },
      ]);
      const reopenedDuplicateCombo = reopened.slides[1]!.shapes.filter(
        (shape): shape is ChartModel => shape instanceof ChartModel,
      ).at(-1)!;
      expect(reopenedDuplicateCombo.definition).toEqual(reopenedCombo.definition);
      const reopenedDuplicatePresentation = reopened.slides[1]!.shapes.filter(
        (shape): shape is ChartModel => shape instanceof ChartModel,
      )[6]!;
      expect(reopenedDuplicatePresentation.definition?.groups[0]?.options?.dataLabels)
        .toMatchObject({
          showCategoryName: true,
          showPercent: true,
          position: 'bestFit',
          numberFormat: '0%',
        });
      for (const chart of charts) {
        expect(reopened.opcPackage.requirePart(chart.workbookPartUri!).contentType)
          .toBe('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        expect(reopened.opcPackage.graph.find(({ uri }) => uri === chart.chartPartUri)?.incoming)
          .toHaveLength(1);
        expect(reopened.opcPackage.graph.find(({ uri }) => uri === chart.workbookPartUri)?.incoming)
          .toHaveLength(1);
      }
    }

    const empty = PptxDocument.create();
    await expect(empty.addChart(0, 'bar', [{
      name: 'Revenue', categories: ['Q1'], values: [1],
    }])).rejects.toThrow(/out of range/);

    if (false) {
      const type: ChartType = 'bar';
      const options: AddChartOptions = { width: inches(6) };
      const groupPromise: Promise<ChartModel> = empty.addChart(0, [{
        type: 'bar',
        series: [{ name: 'Revenue', categories: ['Q1'], values: [1] }],
      }], options);
      // @ts-expect-error chart types are a closed catalog
      const invalidType: ChartType = 'stock';
      // @ts-expect-error chart width uses EMU values
      const invalidOptions: AddChartOptions = { width: '6in' };
      void [type, options, groupPromise, invalidType, invalidOptions];
    }
  }, 15_000);

  it('creates and reopens chart axis foundation options in all six formats', async () => {
    const axisFragment = (xml: string, name: 'catAx' | 'valAx') => {
      const fragment = xml.match(new RegExp(`<c:${name}>[\\s\\S]*?</c:${name}>`, 'u'))?.[0];
      expect(fragment).toBeDefined();
      return fragment!;
    };
    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const document = PptxDocument.create({ format });
      const slide = document.addSlide();
      const chart = await slide.addChart('bar', [{
        name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20],
      }]);
      await chart.replaceDefinition({
        groups: chart.definition!.groups,
        options: {
          categoryAxis: {
            visible: false,
            labelPosition: 'high',
            labelRotation: -45,
            line: { kind: 'none' },
            majorGridLine: {
              kind: 'line',
              color: { kind: 'srgb', value: '445566' },
              width: 0.75,
              dash: 'sysDot',
            },
            majorTickMark: 'inside',
            minorTickMark: 'outside',
          },
          valueAxis: {
            visible: false,
            labelPosition: 'none',
            labelRotation: 45,
            line: {
              kind: 'line',
              color: { kind: 'srgb', value: '778899' },
              width: 1.25,
              dash: 'dash',
            },
            majorGridLine: { kind: 'none' },
            majorTickMark: 'cross',
            minorTickMark: 'none',
          },
        },
      });

      const reopened = await PptxDocument.open(await document.write());
      const reopenedChart = reopened.slides[0]!.shapes.find(
        (shape): shape is ChartModel => shape instanceof ChartModel,
      )!;
      expect(reopenedChart.definition?.options.categoryAxis).toMatchObject({
        visible: false,
        labelPosition: 'high',
        labelRotation: -45,
        line: { kind: 'none' },
        majorGridLine: {
          kind: 'line',
          color: { kind: 'srgb', value: '445566' },
          width: 0.75,
          dash: 'sysDot',
        },
        majorTickMark: 'inside',
        minorTickMark: 'outside',
      });
      expect(reopenedChart.definition?.options.valueAxis).toMatchObject({
        visible: false,
        labelPosition: 'none',
        labelRotation: 45,
        line: {
          kind: 'line',
          color: { kind: 'srgb', value: '778899' },
          width: 1.25,
          dash: 'dash',
        },
        majorGridLine: { kind: 'none' },
        majorTickMark: 'cross',
      });
      expect(reopenedChart.definition?.options.valueAxis?.minorTickMark).toBeUndefined();
      const categoryXml = axisFragment(reopenedChart.xml, 'catAx');
      const valueXml = axisFragment(reopenedChart.xml, 'valAx');
      expect(categoryXml).toContain('<c:delete val="1"/>');
      expect(categoryXml).toContain('<c:tickLblPos val="high"/>');
      expect(categoryXml).toContain('<c:majorTickMark val="in"/>');
      expect(categoryXml).toContain('<c:minorTickMark val="out"/>');
      expect(categoryXml).toContain('rot="-2700000"');
      expect(categoryXml).toContain('<a:ln w="9525"');
      expect(categoryXml).toContain('val="445566"');
      expect(categoryXml).toContain('<a:prstDash val="sysDot"/>');
      expect(valueXml).toContain('<c:delete val="1"/>');
      expect(valueXml).toContain('<c:tickLblPos val="none"/>');
      expect(valueXml).toContain('<c:majorTickMark val="cross"/>');
      expect(valueXml).toContain('<c:minorTickMark val="none"/>');
      expect(valueXml).toContain('rot="2700000"');
      expect(valueXml).toContain('<a:ln w="15875"');
      expect(valueXml).toContain('val="778899"');
      expect(valueXml).toContain('<a:prstDash val="dash"/>');
      expect(validatePackage(reopened.opcPackage).filter(({ severity }) =>
        severity === 'error')).toEqual([]);
    }
  }, 15_000);

  it('edits duplicates rolls back and reopens advanced chart axes in all six formats', async () => {
    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const document = PptxDocument.create({ format });
      const slide = document.addSlide();
      const chart = await slide.addChart('bar3D', [{
        name: 'Revenue',
        categories: [45_658, 45_689],
        values: [1e8, 2e8],
      }]);
      const chartPartUri = chart.chartPartUri;
      if (chartPartUri === undefined) {
        throw new Error('Expected the created chart to have a chart part URI.');
      }
      const initialChartPart = document.opcPackage.requirePart(chartPartUri);
      document.opcPackage.setPart(
        chartPartUri,
        new TextDecoder().decode(initialChartPart.bytes).replace(
          '</c:catAx>',
          '<c:extLst><c:ext uri="urn:axis-keep"/></c:extLst></c:catAx>',
        ),
        initialChartPart.contentType,
      );
      await chart.replaceDefinition({
        groups: chart.definition!.groups,
        options: {
          categoryAxis: {
            kind: 'date',
            crossesAt: 0,
            baseTimeUnit: 'days',
            majorTimeUnit: 'months',
            minorTimeUnit: 'years',
            majorUnit: 2,
            minorUnit: 3,
            labelFrequency: 4,
            multiLevelLabels: true,
          },
          valueAxis: {
            crossesAt: 5,
            displayUnit: 'trillions',
            displayUnitLabel: true,
          },
          seriesAxis: {
            visible: false,
            title: { text: 'Initial depth', color: { kind: 'srgb', value: '112233' } },
            labelPosition: 'low',
            labelFrequency: 3,
            majorUnit: 2,
            minorUnit: 1,
            numberFormat: '0.0',
            orientation: 'maxMin',
            line: { kind: 'none' },
          },
        },
      });
      expect(chart.xml).toContain('uri="urn:axis-keep"');
      const duplicateSlide = document.duplicateSlide(0);
      const duplicateChart = duplicateSlide.shapes.find(
        (shape): shape is ChartModel => shape instanceof ChartModel,
      )!;
      const duplicateDefinition = duplicateChart.definition;

      const beforeInvalidBytes = document.opcPackage.requirePart(chartPartUri).bytes.slice();
      const beforeInvalidJournal = [...document.opcPackage.mutations];
      await expect(chart.replaceDefinition({
        groups: chart.definition!.groups,
        options: {
          categoryAxis: { kind: 'category', baseTimeUnit: 'days' },
        },
      } as never)).rejects.toThrow(/cannot use time units/);
      expect(document.opcPackage.requirePart(chartPartUri).bytes).toEqual(beforeInvalidBytes);
      expect(document.opcPackage.mutations).toEqual(beforeInvalidJournal);

      await chart.replaceDefinition({
        groups: chart.definition!.groups,
        options: {
          categoryAxis: {
            kind: 'date',
            crossesAt: 1,
            baseTimeUnit: 'months',
            majorTimeUnit: 'years',
            majorUnit: 1,
            labelFrequency: 2,
            multiLevelLabels: true,
          },
          valueAxis: {
            crossesAt: 10,
            displayUnit: 'hundredMillions',
            displayUnitLabel: true,
          },
          seriesAxis: {
            visible: false,
            title: { text: 'Edited depth', color: { kind: 'srgb', value: '445566' } },
            labelPosition: 'high',
            labelFrequency: 2,
            majorUnit: 4,
            minorUnit: 2,
            numberFormat: '#,##0',
            line: { kind: 'line', color: { kind: 'srgb', value: '778899' } },
          },
        },
      });
      expect(duplicateChart.definition).toEqual(duplicateDefinition);

      const reopened = await PptxDocument.open(await document.write());
      const source = reopened.slides[0]!.shapes.find(
        (shape): shape is ChartModel => shape instanceof ChartModel,
      )!;
      const duplicate = reopened.slides[1]!.shapes.find(
        (shape): shape is ChartModel => shape instanceof ChartModel,
      )!;
      expect(source.definition?.options.categoryAxis).toMatchObject({
        kind: 'date',
        crossesAt: 1,
        baseTimeUnit: 'months',
        majorTimeUnit: 'years',
        majorUnit: 1,
        labelFrequency: 2,
        multiLevelLabels: true,
      });
      expect(source.definition?.options.valueAxis).toMatchObject({
        crossesAt: 10,
        displayUnit: 'hundredMillions',
        displayUnitLabel: true,
      });
      expect(source.definition?.options.seriesAxis).toMatchObject({
        visible: false,
        title: { text: 'Edited depth', color: { kind: 'srgb', value: '445566' } },
        labelPosition: 'high',
        labelFrequency: 2,
        majorUnit: 4,
        minorUnit: 2,
        numberFormat: '#,##0',
        line: { kind: 'line', color: { kind: 'srgb', value: '778899' } },
      });
      expect(duplicate.definition).toEqual(duplicateDefinition);
      expect(source.xml).toContain('<c:dateAx>');
      expect(source.xml).toContain('uri="urn:axis-keep"');
      expect(source.xml).toContain('<c:builtInUnit val="hundredMillions"/>');
      expect(source.xml).toContain('<a:t>Edited depth</a:t>');
      expect(validatePackage(reopened.opcPackage).filter(({ severity }) =>
        severity === 'error')).toEqual([]);
    }
  }, 30_000);

  it('validates chart caches and workbooks on strict writes without mutating the package', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const chart = await document.addChart(0, 'bar', [{
      name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20],
    }]);
    await document.write();
    expect(document.diagnostics.filter(({ code }) => code.startsWith('CHART_'))).toEqual([]);

    chart.setXml(chart.xml.replace('<c:v>20</c:v>', '<c:v>21</c:v>'));
    const before = new Map(document.opcPackage.parts.map(({ uri, bytes }) => [uri, bytes.slice()]));
    const journal = [...document.opcPackage.mutations];
    await expect(document.write()).rejects.toBeInstanceOf(ValidationError);
    expect(document.diagnostics).toMatchObject([{
      severity: 'error',
      code: 'CHART_WORKBOOK_CACHE_DIVERGENCE',
      partUri: chart.chartPartUri,
      objectId: String(chart.id),
    }]);
    expect(document.opcPackage.parts.map(({ uri }) => uri)).toEqual([...before.keys()]);
    for (const { uri, bytes } of document.opcPackage.parts) expect(bytes).toEqual(before.get(uri));
    expect(document.opcPackage.mutations).toEqual(journal);
    await expect(document.write({ mode: 'permissive' })).resolves.toBeInstanceOf(Uint8Array);

    const cacheOnly = PptxDocument.create();
    const cacheOnlySlide = cacheOnly.addSlide();
    const cacheOnlyChart = await cacheOnly.addChart(0, 'line', [{
      name: 'Trend', categories: ['Q1'], values: [1],
    }]);
    cacheOnlyChart.setXml(
      cacheOnlyChart.xml.replace(/<c:externalData[\s\S]*?<\/c:externalData>/, ''),
    );
    await expect(cacheOnly.write()).resolves.toBeInstanceOf(Uint8Array);
    expect(cacheOnly.diagnostics).toMatchObject([{
      severity: 'warning',
      code: 'CHART_WORKBOOK_MISSING',
      partUri: cacheOnlyChart.chartPartUri,
      objectId: String(cacheOnlyChart.id),
    }]);
    expect(cacheOnlySlide.shapes[0]).toBe(cacheOnlyChart);
  });

  it('creates all presentation formats and built-in PptxGenJS slide sizes', async () => {
    const sizes = {
      '4:3': [9_144_000, 6_858_000],
      '16:9': [9_144_000, 5_143_500],
      '16:10': [9_144_000, 5_715_000],
      wide: [12_192_000, 6_858_000],
    } as const;
    for (const [slideSize, [cx, cy]] of Object.entries(sizes)) {
      const document = PptxDocument.create({ slideSize: slideSize as keyof typeof sizes });
      expect(new TextDecoder().decode(document.opcPackage.requirePart('/ppt/presentation.xml').bytes)).toContain(
        `<p:sldSz cx="${cx}" cy="${cy}"/>`,
      );
    }

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({ format });
      expect(created.formatProfile).toEqual(PRESENTATION_FORMAT_PROFILES[format]);
      const reopened = await PptxDocument.open(await created.write());
      expect(reopened.formatProfile).toEqual(PRESENTATION_FORMAT_PROFILES[format]);
    }
  });

  it('creates and round-trips an OOXML-valid custom slide size without retaining its input', async () => {
    const slideSize = { width: inches(11.7), height: inches(8.3) };
    const document = PptxDocument.create({ slideSize });
    slideSize.width = inches(10);
    expect(document.slideSize).toEqual({ width: inches(11.7), height: inches(8.3) });
    const presentationXml = new TextDecoder().decode(
      document.opcPackage.requirePart('/ppt/presentation.xml').bytes,
    );
    expect(presentationXml).toContain('<p:sldSz cx="10698480" cy="7589520"/>');
    expect(presentationXml).toContain('<p:notesSz cx="7589520" cy="10698480"/>');

    document.addSlide();
    const reopened = await PptxDocument.open(await document.write());
    reopened.slideSize = { width: inches(10), height: inches(7.5) };
    expect(reopened.slideSize).toEqual({ width: inches(10), height: inches(7.5) });
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
    expect(new TextDecoder().decode(reopened.opcPackage.requirePart('/ppt/presentation.xml').bytes)).toContain(
      '<p:sldSz cx="9144000" cy="6858000"/>',
    );
  });

  it('reads, replaces, rolls back, and reopens the presentation-direct theme fonts', async () => {
    const document = PptxDocument.create();
    const theme = document.masterLayoutTheme.presentationTheme!;
    const initial = document.theme!;
    expect(initial).toEqual({ headFontFace: 'Aptos Display', bodyFontFace: 'Aptos' });
    expect(document.theme).not.toBe(initial);
    expect(document.masterLayoutTheme.presentationTheme).toBe(theme);
    (initial as { headFontFace: string }).headFontFace = 'Changed detached snapshot';
    expect(document.theme).toEqual({ headFontFace: 'Aptos Display', bodyFontFace: 'Aptos' });

    const unrelated = new Map(
      document.opcPackage.parts
        .filter(({ uri }) => uri !== theme.partUri)
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    document.theme = { headFontFace: 'Noto Sans Display' };
    expect(document.theme).toEqual({
      headFontFace: 'Noto Sans Display',
      bodyFontFace: 'Calibri',
    });
    for (const [uri, bytes] of unrelated) {
      expect(document.opcPackage.requirePart(uri).bytes).toEqual(bytes);
    }

    document.theme = { bodyFontFace: 'Noto Sans' };
    expect(document.theme).toEqual({
      headFontFace: 'Calibri Light',
      bodyFontFace: 'Noto Sans',
    });
    document.theme = {};
    expect(document.theme).toEqual({
      headFontFace: 'Calibri Light',
      bodyFontFace: 'Calibri',
    });
    document.theme = {
      headFontFace: 'A&B <Display> "One"',
      bodyFontFace: 'A&B <Body> "Two"',
    };
    expect(document.theme).toEqual({
      headFontFace: 'A&B <Display> "One"',
      bodyFontFace: 'A&B <Body> "Two"',
    });
    const themeXml = new TextDecoder().decode(
      document.opcPackage.requirePart(theme.partUri).bytes,
    );
    expect(themeXml).toContain('typeface="A&amp;B &lt;Display&gt; &quot;One&quot;"');
    expect(themeXml).toContain('typeface="A&amp;B &lt;Body&gt; &quot;Two&quot;"');

    const beforeNoOp = document.opcPackage.requirePart(theme.partUri).bytes;
    const noOpJournal = [...document.opcPackage.mutations];
    document.theme = {
      headFontFace: 'A&B <Display> "One"',
      bodyFontFace: 'A&B <Body> "Two"',
    };
    expect(document.opcPackage.requirePart(theme.partUri).bytes).toEqual(beforeNoOp);
    expect(document.opcPackage.mutations).toEqual(noOpJournal);

    expect(() =>
      document.transaction(() => {
        document.theme = {
          headFontFace: 'Temporary Display',
          bodyFontFace: 'Temporary Body',
        };
        throw new Error('restore presentation theme fonts');
      }),
    ).toThrow('restore presentation theme fonts');
    expect(document.opcPackage.requirePart(theme.partUri).bytes).toEqual(beforeNoOp);
    expect(document.opcPackage.mutations).toEqual(noOpJournal);

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.theme).toEqual({
      headFontFace: 'A&B <Display> "One"',
      bodyFontFace: 'A&B <Body> "Two"',
    });
    expect(reopened.masterLayoutTheme.presentationTheme?.fonts).toEqual({
      majorLatin: 'A&B <Display> "One"',
      minorLatin: 'A&B <Body> "Two"',
    });
  });

  it('rejects malformed live theme values and unsafe primary relationships without mutation', () => {
    const document = PptxDocument.create();
    const themePartUri = document.masterLayoutTheme.presentationTheme!.partUri;
    const before = document.opcPackage.requirePart(themePartUri).bytes;
    const journal = [...document.opcPackage.mutations];
    const accessor = Object.create(null) as Record<string, unknown>;
    let getterCalls = 0;
    Object.defineProperty(accessor, 'headFontFace', {
      enumerable: true,
      get() {
        getterCalls += 1;
        return 'Never';
      },
    });
    const symbol = { headFontFace: 'Aptos Display', [Symbol('extra')]: true };
    const customPrototype = Object.create({ inherited: true }) as { headFontFace?: string };
    customPrototype.headFontFace = 'Aptos Display';
    for (const value of [
      undefined,
      null,
      [],
      'Aptos',
      { unknown: 'Aptos' },
      { headFontFace: '' },
      { headFontFace: '   ' },
      { headFontFace: 1 },
      { bodyFontFace: 'bad\u0000font' },
      accessor,
      symbol,
      customPrototype,
    ]) {
      expect(() => {
        document.theme = value as never;
      }).toThrow(TypeError);
    }
    expect(getterCalls).toBe(0);
    expect(document.opcPackage.requirePart(themePartUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);

    const noDirect = PptxDocument.create();
    noDirect.opcPackage.removeRelationship('/ppt/presentation.xml', 'rId6');
    expect(noDirect.theme).toBeUndefined();
    const noDirectJournal = [...noDirect.opcPackage.mutations];
    expect(() => {
      noDirect.theme = { headFontFace: 'Noto Sans Display' };
    }).toThrow(/one editable direct theme/);
    expect(noDirect.opcPackage.mutations).toEqual(noDirectJournal);

    const duplicate = PptxDocument.create();
    duplicate.opcPackage.addRelationship('/ppt/presentation.xml', {
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
      target: 'theme/theme1.xml',
    });
    expect(duplicate.theme).toBeUndefined();
    const duplicateJournal = [...duplicate.opcPackage.mutations];
    expect(() => {
      duplicate.theme = { bodyFontFace: 'Noto Sans' };
    }).toThrow(/one editable direct theme/);
    expect(duplicate.opcPackage.mutations).toEqual(duplicateJournal);

    const external = PptxDocument.create();
    external.opcPackage.updateRelationship('/ppt/presentation.xml', 'rId6', {
      target: 'https://example.com/theme.xml',
      targetMode: 'External',
    });
    expect(external.theme).toBeUndefined();
    const externalJournal = [...external.opcPackage.mutations];
    expect(() => {
      external.theme = {};
    }).toThrow(/one editable direct theme/);
    expect(external.opcPackage.mutations).toEqual(externalJournal);

    const dangling = PptxDocument.create();
    const relationshipPart = dangling.opcPackage.requirePart('/ppt/_rels/presentation.xml.rels');
    dangling.opcPackage.setPart(
      relationshipPart.uri,
      new TextDecoder().decode(relationshipPart.bytes).replace(
        'Target="theme/theme1.xml"',
        'Target="theme/missing.xml"',
      ),
      relationshipPart.contentType,
    );
    expect(dangling.theme).toBeUndefined();
    const danglingJournal = [...dangling.opcPackage.mutations];
    expect(() => {
      dangling.theme = {};
    }).toThrow(/one editable direct theme/);
    expect(dangling.opcPackage.mutations).toEqual(danglingJournal);

    const wrongType = PptxDocument.create();
    const wrongTypeTheme = wrongType.opcPackage.requirePart('/ppt/theme/theme1.xml');
    wrongType.opcPackage.setPart(
      wrongTypeTheme.uri,
      wrongTypeTheme.bytes,
      'application/xml',
    );
    expect(wrongType.theme).toBeUndefined();
    const wrongTypeJournal = [...wrongType.opcPackage.mutations];
    expect(() => {
      wrongType.theme = {};
    }).toThrow(/one editable direct theme/);
    expect(wrongType.opcPackage.mutations).toEqual(wrongTypeJournal);
  });

  it('creates presentation theme fonts with explicit PptxGenJS partial fallbacks', async () => {
    const omitted = PptxDocument.create();
    const explicitUndefined = PptxDocument.create({ theme: undefined } as never);
    const empty = PptxDocument.create({ theme: {} });
    const headOnly = PptxDocument.create({
      theme: { headFontFace: 'Noto Sans Display' },
    });
    const bodyOnly = PptxDocument.create({
      theme: { bodyFontFace: 'Noto Sans' },
    });
    const both = PptxDocument.create({
      theme: {
        headFontFace: 'Aptos Display',
        bodyFontFace: 'Aptos',
      },
    });

    expect(omitted.theme).toEqual({ headFontFace: 'Aptos Display', bodyFontFace: 'Aptos' });
    expect(explicitUndefined.theme).toEqual(omitted.theme);
    expect(await explicitUndefined.write()).toEqual(await omitted.write());
    expect(empty.theme).toEqual({ headFontFace: 'Calibri Light', bodyFontFace: 'Calibri' });
    expect(headOnly.theme).toEqual({
      headFontFace: 'Noto Sans Display',
      bodyFontFace: 'Calibri',
    });
    expect(bodyOnly.theme).toEqual({
      headFontFace: 'Calibri Light',
      bodyFontFace: 'Noto Sans',
    });
    expect(both.theme).toEqual({ headFontFace: 'Aptos Display', bodyFontFace: 'Aptos' });

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({
        format,
        theme: { headFontFace: `Display ${format}`, bodyFontFace: `Body ${format}` },
      });
      expect(created.theme).toEqual({
        headFontFace: `Display ${format}`,
        bodyFontFace: `Body ${format}`,
      });
      expect((await PptxDocument.open(await created.write())).theme).toEqual(created.theme);
    }

    const input = {
      headFontFace: 'Detached Display',
      bodyFontFace: 'Detached Body',
    };
    const detached = PptxDocument.create({ theme: input });
    input.headFontFace = 'Changed Display';
    input.bodyFontFace = 'Changed Body';
    expect(detached.theme).toEqual({
      headFontFace: 'Detached Display',
      bodyFontFace: 'Detached Body',
    });
    expect(PptxDocument.create({
      theme: Object.freeze({ headFontFace: 'Frozen Display' }),
    }).theme).toEqual({ headFontFace: 'Frozen Display', bodyFontFace: 'Calibri' });
    const nullPrototype = Object.assign(Object.create(null), {
      bodyFontFace: 'Null Prototype Body',
    });
    expect(PptxDocument.create({ theme: nullPrototype }).theme).toEqual({
      headFontFace: 'Calibri Light',
      bodyFontFace: 'Null Prototype Body',
    });
    const escaped = PptxDocument.create({
      theme: {
        headFontFace: 'A&B <Display> "One"',
        bodyFontFace: 'A&B <Body> "Two"',
      },
    });
    expect(escaped.theme).toEqual({
      headFontFace: 'A&B <Display> "One"',
      bodyFontFace: 'A&B <Body> "Two"',
    });
    expect(new TextDecoder().decode(
      escaped.opcPackage.requirePart('/ppt/theme/theme1.xml').bytes,
    )).toContain('typeface="A&amp;B &lt;Display&gt; &quot;One&quot;"');
  });

  it('rejects malformed presentation theme creation input before returning a document', () => {
    const accessor = Object.create(null) as Record<string, unknown>;
    let getterCalls = 0;
    Object.defineProperty(accessor, 'headFontFace', {
      enumerable: true,
      get() {
        getterCalls += 1;
        return 'Never';
      },
    });
    const symbol = { headFontFace: 'Aptos Display', [Symbol('extra')]: true };
    const customPrototype = Object.create({ inherited: true }) as { headFontFace?: string };
    customPrototype.headFontFace = 'Aptos Display';
    for (const theme of [
      null,
      [],
      'Aptos',
      { unknown: 'Aptos' },
      { headFontFace: '' },
      { headFontFace: '   ' },
      { bodyFontFace: 1 },
      { bodyFontFace: 'bad\u0000font' },
      accessor,
      symbol,
      customPrototype,
    ]) {
      expect(() => PptxDocument.create({ theme: theme as never })).toThrow(TypeError);
    }
    expect(getterCalls).toBe(0);
  });

  it('creates, edits, clears, rolls back, and reopens presentation title metadata', async () => {
    const readCoreXml = (document: PptxDocument): string => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    const omitted = PptxDocument.create();
    const explicitUndefined = PptxDocument.create({ title: undefined } as never);
    const custom = PptxDocument.create({ title: 'Quarterly & <Review>' });
    const empty = PptxDocument.create({ title: '' });

    expect([omitted.title, explicitUndefined.title, custom.title, empty.title]).toEqual([
      undefined,
      undefined,
      'Quarterly & <Review>',
      '',
    ]);
    expect(readCoreXml(omitted)).not.toContain('<dc:title');
    expect(readCoreXml(explicitUndefined)).not.toContain('<dc:title');
    expect(readCoreXml(custom)).toContain(
      '<dc:title>Quarterly &amp; &lt;Review&gt;</dc:title>',
    );
    expect(readCoreXml(empty)).toContain('<dc:title></dc:title>');

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({ format, title: `Title ${format}` });
      expect(created.title).toBe(`Title ${format}`);
      expect(validatePackage(created.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);
      const reopened = await PptxDocument.open(await created.write());
      expect(reopened.title).toBe(`Title ${format}`);
      expect(reopened.format).toBe(format);
    }

    const beforeSame = custom.opcPackage.requirePart('/docProps/core.xml').bytes;
    const sameJournal = [...custom.opcPackage.mutations];
    custom.title = 'Quarterly & <Review>';
    expect(custom.opcPackage.requirePart('/docProps/core.xml').bytes).toEqual(beforeSame);
    expect(custom.opcPackage.mutations).toEqual(sameJournal);

    const slide = custom.addSlide();
    custom.title = 'Edited title';
    expect(custom.title).toBe('Edited title');
    expect(custom.slides[0]).toBe(slide);

    const beforeRollback = custom.opcPackage.requirePart('/docProps/core.xml').bytes;
    const rollbackJournal = [...custom.opcPackage.mutations];
    expect(() => custom.transaction(() => {
      custom.title = 'Temporary title';
      expect(custom.title).toBe('Temporary title');
      throw new Error('restore presentation title');
    })).toThrow('restore presentation title');
    expect(custom.opcPackage.requirePart('/docProps/core.xml').bytes).toEqual(beforeRollback);
    expect(custom.opcPackage.mutations).toEqual(rollbackJournal);
    expect(custom.title).toBe('Edited title');
    expect(custom.slides[0]).toBe(slide);

    const reopenedEdited = await PptxDocument.open(await custom.write());
    expect(reopenedEdited.title).toBe('Edited title');
    reopenedEdited.title = '';
    expect(reopenedEdited.title).toBe('');
    reopenedEdited.title = undefined;
    expect(reopenedEdited.title).toBeUndefined();
    expect(readCoreXml(reopenedEdited)).not.toContain('<dc:title');
    expect(readCoreXml(reopenedEdited)).toContain('<dc:creator>@jiayunxie/pptx</dc:creator>');
    expect(readCoreXml(reopenedEdited)).toContain('<cp:revision>1</cp:revision>');
    const reopenedCleared = await PptxDocument.open(await reopenedEdited.write());
    expect(reopenedCleared.title).toBeUndefined();
    expect(reopenedCleared.slides).toHaveLength(1);
  });

  it('rejects malformed presentation title metadata during creation', () => {
    for (const title of [
      null,
      true,
      false,
      0,
      1,
      {},
      [],
      Symbol('title'),
      'bad\u0001title',
    ]) {
      expect(() => PptxDocument.create({ title: title as never })).toThrow(TypeError);
    }
  });

  it('creates, edits, clears, rolls back, and reopens presentation author metadata', async () => {
    const readCoreXml = (document: PptxDocument): string => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    const omitted = PptxDocument.create();
    const explicitUndefined = PptxDocument.create({ author: undefined } as never);
    const custom = PptxDocument.create({ author: 'Alice & <Bob>' });
    const empty = PptxDocument.create({ author: '' });

    expect([omitted.author, explicitUndefined.author, custom.author, empty.author]).toEqual([
      '@jiayunxie/pptx',
      '@jiayunxie/pptx',
      'Alice & <Bob>',
      '',
    ]);
    expect(readCoreXml(explicitUndefined)).toBe(readCoreXml(omitted));
    expect(readCoreXml(custom)).toContain(
      '<dc:creator>Alice &amp; &lt;Bob&gt;</dc:creator>',
    );
    expect(readCoreXml(custom)).toContain(
      '<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>',
    );
    expect(readCoreXml(empty)).toContain('<dc:creator></dc:creator>');
    expect(readCoreXml(empty)).toContain(
      '<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>',
    );

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({ format, author: `Author ${format}` });
      expect(created.author).toBe(`Author ${format}`);
      expect(readCoreXml(created)).toContain(
        '<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>',
      );
      expect(validatePackage(created.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);
      const reopened = await PptxDocument.open(await created.write());
      expect(reopened.author).toBe(`Author ${format}`);
      expect(reopened.format).toBe(format);
    }

    const combined = PptxDocument.create({
      author: 'Combined author',
      title: 'Combined title',
    });
    expect([combined.author, combined.title]).toEqual(['Combined author', 'Combined title']);
    expect(readCoreXml(combined)).toContain('<dc:creator>Combined author</dc:creator>');
    expect(readCoreXml(combined)).toContain('<dc:title>Combined title</dc:title>');
    expect(readCoreXml(combined)).toContain(
      '<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>',
    );

    const beforeSame = custom.opcPackage.requirePart('/docProps/core.xml').bytes;
    const sameJournal = [...custom.opcPackage.mutations];
    custom.author = 'Alice & <Bob>';
    expect(custom.opcPackage.requirePart('/docProps/core.xml').bytes).toEqual(beforeSame);
    expect(custom.opcPackage.mutations).toEqual(sameJournal);

    const slide = custom.addSlide();
    const otherParts = new Map(
      custom.opcPackage.parts
        .filter(({ uri }) => uri !== '/docProps/core.xml')
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    custom.author = 'Edited author';
    expect(custom.author).toBe('Edited author');
    expect(custom.slides[0]).toBe(slide);
    expect(readCoreXml(custom)).toContain('<dc:creator>Edited author</dc:creator>');
    expect(readCoreXml(custom)).toContain(
      '<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>',
    );
    for (const [uri, bytes] of otherParts) {
      expect(custom.opcPackage.requirePart(uri).bytes).toEqual(bytes);
    }

    const beforeRollback = custom.opcPackage.requirePart('/docProps/core.xml').bytes;
    const rollbackJournal = [...custom.opcPackage.mutations];
    expect(() => custom.transaction(() => {
      custom.author = 'Temporary author';
      expect(custom.author).toBe('Temporary author');
      throw new Error('restore presentation author');
    })).toThrow('restore presentation author');
    expect(custom.opcPackage.requirePart('/docProps/core.xml').bytes).toEqual(beforeRollback);
    expect(custom.opcPackage.mutations).toEqual(rollbackJournal);
    expect(custom.author).toBe('Edited author');
    expect(custom.slides[0]).toBe(slide);

    const reopenedEdited = await PptxDocument.open(await custom.write());
    expect(reopenedEdited.author).toBe('Edited author');
    reopenedEdited.author = '';
    expect(reopenedEdited.author).toBe('');
    reopenedEdited.author = undefined;
    expect(reopenedEdited.author).toBeUndefined();
    expect(readCoreXml(reopenedEdited)).not.toContain('<dc:creator');
    expect(readCoreXml(reopenedEdited)).toContain(
      '<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>',
    );
    expect(readCoreXml(reopenedEdited)).toContain('<cp:revision>1</cp:revision>');
    const reopenedCleared = await PptxDocument.open(await reopenedEdited.write());
    expect(reopenedCleared.author).toBeUndefined();
    expect(reopenedCleared.slides).toHaveLength(1);
  });

  it('rejects malformed presentation author metadata during creation', () => {
    for (const author of [
      null,
      true,
      false,
      0,
      1,
      {},
      [],
      Symbol('author'),
      'bad\u0001author',
    ]) {
      expect(() => PptxDocument.create({ author: author as never })).toThrow(TypeError);
    }
  });

  it('creates, edits, clears, rolls back, and reopens presentation last modified by metadata', async () => {
    const readCoreXml = (document: PptxDocument): string => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    const omitted = PptxDocument.create();
    const explicitUndefined = PptxDocument.create({ lastModifiedBy: undefined } as never);
    const custom = PptxDocument.create({ lastModifiedBy: 'Editor & <Reviewer>' });
    const empty = PptxDocument.create({ lastModifiedBy: '' });

    expect([
      omitted.lastModifiedBy,
      explicitUndefined.lastModifiedBy,
      custom.lastModifiedBy,
      empty.lastModifiedBy,
    ]).toEqual([
      '@jiayunxie/pptx',
      '@jiayunxie/pptx',
      'Editor & <Reviewer>',
      '',
    ]);
    expect(readCoreXml(explicitUndefined)).toBe(readCoreXml(omitted));
    expect(readCoreXml(custom)).toContain(
      '<cp:lastModifiedBy>Editor &amp; &lt;Reviewer&gt;</cp:lastModifiedBy>',
    );
    expect(readCoreXml(custom)).toContain(
      '<dc:creator>@jiayunxie/pptx</dc:creator>',
    );
    expect(readCoreXml(empty)).toContain('<cp:lastModifiedBy></cp:lastModifiedBy>');
    expect(readCoreXml(empty)).toContain(
      '<dc:creator>@jiayunxie/pptx</dc:creator>',
    );

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({
        format,
        lastModifiedBy: `Editor ${format}`,
      });
      expect(created.lastModifiedBy).toBe(`Editor ${format}`);
      expect(readCoreXml(created)).toContain(
        `<cp:lastModifiedBy>Editor ${format}</cp:lastModifiedBy>`,
      );
      expect(validatePackage(created.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);
      const reopened = await PptxDocument.open(await created.write());
      expect(reopened.lastModifiedBy).toBe(`Editor ${format}`);
      expect(reopened.format).toBe(format);
    }

    const combined = PptxDocument.create({
      author: 'Combined author',
      lastModifiedBy: 'Combined editor',
      revision: '8',
      subject: 'Combined subject',
      title: 'Combined title',
    });
    expect([
      combined.author,
      combined.lastModifiedBy,
      combined.revision,
      combined.subject,
      combined.title,
    ]).toEqual([
      'Combined author',
      'Combined editor',
      '8',
      'Combined subject',
      'Combined title',
    ]);
    expect(readCoreXml(combined)).toContain('<dc:creator>Combined author</dc:creator>');
    expect(readCoreXml(combined)).toContain(
      '<cp:lastModifiedBy>Combined editor</cp:lastModifiedBy>',
    );
    expect(readCoreXml(combined)).toContain('<cp:revision>8</cp:revision>');

    const beforeSame = custom.opcPackage.requirePart('/docProps/core.xml').bytes;
    const sameJournal = [...custom.opcPackage.mutations];
    custom.lastModifiedBy = 'Editor & <Reviewer>';
    expect(custom.opcPackage.requirePart('/docProps/core.xml').bytes).toEqual(beforeSame);
    expect(custom.opcPackage.mutations).toEqual(sameJournal);

    const slide = custom.addSlide();
    const otherParts = new Map(
      custom.opcPackage.parts
        .filter(({ uri }) => uri !== '/docProps/core.xml')
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    custom.lastModifiedBy = 'Edited editor';
    expect(custom.lastModifiedBy).toBe('Edited editor');
    expect(custom.author).toBe('@jiayunxie/pptx');
    expect(custom.slides[0]).toBe(slide);
    expect(readCoreXml(custom)).toContain(
      '<cp:lastModifiedBy>Edited editor</cp:lastModifiedBy>',
    );
    expect(readCoreXml(custom)).toContain(
      '<dc:creator>@jiayunxie/pptx</dc:creator>',
    );
    for (const [uri, bytes] of otherParts) {
      expect(custom.opcPackage.requirePart(uri).bytes).toEqual(bytes);
    }

    custom.author = 'Edited author';
    expect(custom.lastModifiedBy).toBe('Edited editor');
    expect(custom.author).toBe('Edited author');

    const beforeRollback = custom.opcPackage.requirePart('/docProps/core.xml').bytes;
    const rollbackJournal = [...custom.opcPackage.mutations];
    expect(() => custom.transaction(() => {
      custom.lastModifiedBy = 'Temporary editor';
      expect(custom.lastModifiedBy).toBe('Temporary editor');
      throw new Error('restore presentation lastModifiedBy');
    })).toThrow('restore presentation lastModifiedBy');
    expect(custom.opcPackage.requirePart('/docProps/core.xml').bytes).toEqual(beforeRollback);
    expect(custom.opcPackage.mutations).toEqual(rollbackJournal);
    expect(custom.lastModifiedBy).toBe('Edited editor');
    expect(custom.slides[0]).toBe(slide);

    const reopenedEdited = await PptxDocument.open(await custom.write());
    expect(reopenedEdited.lastModifiedBy).toBe('Edited editor');
    expect(reopenedEdited.author).toBe('Edited author');
    reopenedEdited.lastModifiedBy = '';
    expect(reopenedEdited.lastModifiedBy).toBe('');
    expect(reopenedEdited.author).toBe('Edited author');
    reopenedEdited.lastModifiedBy = undefined;
    expect(reopenedEdited.lastModifiedBy).toBeUndefined();
    expect(readCoreXml(reopenedEdited)).not.toContain('<cp:lastModifiedBy');
    expect(readCoreXml(reopenedEdited)).toContain('<dc:creator>Edited author</dc:creator>');
    expect(readCoreXml(reopenedEdited)).toContain('<cp:revision>1</cp:revision>');
    const reopenedCleared = await PptxDocument.open(await reopenedEdited.write());
    expect(reopenedCleared.lastModifiedBy).toBeUndefined();
    expect(reopenedCleared.author).toBe('Edited author');
    expect(reopenedCleared.slides).toHaveLength(1);
  });

  it('rejects malformed presentation last modified by metadata during creation', () => {
    for (const lastModifiedBy of [
      null,
      true,
      false,
      0,
      1,
      1n,
      {},
      [],
      Symbol('lastModifiedBy'),
      'bad\u0001editor',
    ]) {
      expect(() => PptxDocument.create({ lastModifiedBy } as never)).toThrow(TypeError);
    }
  });

  it('creates, edits, clears, rolls back, and reopens presentation created-at metadata', async () => {
    const readCoreXml = (document: PptxDocument): string => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    const omitted = PptxDocument.create();
    const explicitUndefined = PptxDocument.create({ createdAt: undefined } as never);
    const custom = PptxDocument.create({
      createdAt: '2024-02-29T12:34:56.123456+05:30',
    });

    expect([omitted.createdAt, explicitUndefined.createdAt, custom.createdAt]).toEqual([
      undefined,
      undefined,
      '2024-02-29T12:34:56.123456+05:30',
    ]);
    expect(readCoreXml(explicitUndefined)).toBe(readCoreXml(omitted));
    expect(readCoreXml(omitted)).not.toContain('<dcterms:created');
    expect(readCoreXml(custom)).toMatch(
      /<dcterms:created\b[^>]*\bxsi:type="dcterms:W3CDTF"[^>]*>2024-02-29T12:34:56\.123456\+05:30<\/dcterms:created>/,
    );
    expect(readCoreXml(custom).match(/xmlns:dcterms=/g)).toHaveLength(1);
    expect(readCoreXml(custom).match(/xmlns:xsi=/g)).toHaveLength(1);

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({
        createdAt: '2026-07-30T00:00:00Z',
        format,
      });
      expect(created.createdAt).toBe('2026-07-30T00:00:00Z');
      expect(readCoreXml(created)).toMatch(
        /<dcterms:created\b[^>]*\bxsi:type="dcterms:W3CDTF"[^>]*>2026-07-30T00:00:00Z<\/dcterms:created>/,
      );
      expect(validatePackage(created.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);
      const reopened = await PptxDocument.open(await created.write());
      expect(reopened.createdAt).toBe('2026-07-30T00:00:00Z');
      expect(reopened.format).toBe(format);
    }

    const combined = PptxDocument.create({
      author: 'Combined author',
      company: 'Combined company',
      createdAt: '2024-02-29T12:34:56.123+05:30',
      lastModifiedBy: 'Combined editor',
      revision: '8',
      subject: 'Combined subject',
      title: 'Combined title',
    });
    expect([
      combined.author,
      combined.company,
      combined.createdAt,
      combined.lastModifiedBy,
      combined.revision,
      combined.subject,
      combined.title,
    ]).toEqual([
      'Combined author',
      'Combined company',
      '2024-02-29T12:34:56.123+05:30',
      'Combined editor',
      '8',
      'Combined subject',
      'Combined title',
    ]);
    expect(readCoreXml(combined)).toContain('<dc:creator>Combined author</dc:creator>');
    expect(readCoreXml(combined)).toContain(
      '<cp:lastModifiedBy>Combined editor</cp:lastModifiedBy>',
    );
    expect(readCoreXml(combined)).toContain('<cp:revision>8</cp:revision>');
    expect(readCoreXml(combined)).toContain('<dc:subject>Combined subject</dc:subject>');
    expect(readCoreXml(combined)).toContain('<dc:title>Combined title</dc:title>');
    expect(new TextDecoder().decode(
      combined.opcPackage.requirePart('/docProps/app.xml').bytes,
    )).toContain('<Company>Combined company</Company>');

    const beforeSame = custom.opcPackage.requirePart('/docProps/core.xml').bytes;
    const sameJournal = [...custom.opcPackage.mutations];
    custom.createdAt = '2024-02-29T12:34:56.123456+05:30';
    expect(custom.opcPackage.requirePart('/docProps/core.xml').bytes).toEqual(beforeSame);
    expect(custom.opcPackage.mutations).toEqual(sameJournal);

    const slide = custom.addSlide();
    const otherParts = new Map(
      custom.opcPackage.parts
        .filter(({ uri }) => uri !== '/docProps/core.xml')
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    custom.createdAt = '2026-07-30T00:00:00Z';
    expect(custom.createdAt).toBe('2026-07-30T00:00:00Z');
    expect(custom.author).toBe('@jiayunxie/pptx');
    expect(custom.lastModifiedBy).toBe('@jiayunxie/pptx');
    expect(custom.revision).toBe('1');
    expect(custom.slides[0]).toBe(slide);
    expect(readCoreXml(custom)).toMatch(
      /<dcterms:created\b[^>]*\bxsi:type="dcterms:W3CDTF"[^>]*>2026-07-30T00:00:00Z<\/dcterms:created>/,
    );
    for (const [uri, bytes] of otherParts) {
      expect(custom.opcPackage.requirePart(uri).bytes).toEqual(bytes);
    }

    const beforeRollback = custom.opcPackage.requirePart('/docProps/core.xml').bytes;
    const rollbackJournal = [...custom.opcPackage.mutations];
    expect(() => custom.transaction(() => {
      custom.createdAt = '2025-01-01T00:00:00Z';
      expect(custom.createdAt).toBe('2025-01-01T00:00:00Z');
      throw new Error('restore presentation createdAt');
    })).toThrow('restore presentation createdAt');
    expect(custom.opcPackage.requirePart('/docProps/core.xml').bytes).toEqual(beforeRollback);
    expect(custom.opcPackage.mutations).toEqual(rollbackJournal);
    expect(custom.createdAt).toBe('2026-07-30T00:00:00Z');
    expect(custom.slides[0]).toBe(slide);

    const reopenedEdited = await PptxDocument.open(await custom.write());
    expect(reopenedEdited.createdAt).toBe('2026-07-30T00:00:00Z');
    reopenedEdited.createdAt = undefined;
    expect(reopenedEdited.createdAt).toBeUndefined();
    expect(readCoreXml(reopenedEdited)).not.toContain('<dcterms:created');
    expect(readCoreXml(reopenedEdited)).toContain(
      '<dc:creator>@jiayunxie/pptx</dc:creator>',
    );
    expect(readCoreXml(reopenedEdited)).toContain(
      '<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>',
    );
    expect(readCoreXml(reopenedEdited)).toContain('<cp:revision>1</cp:revision>');
    const reopenedCleared = await PptxDocument.open(await reopenedEdited.write());
    expect(reopenedCleared.createdAt).toBeUndefined();
    expect(reopenedCleared.slides).toHaveLength(1);
  });

  it('rejects malformed presentation created-at metadata during creation', () => {
    for (const createdAt of [
      '',
      ' 2026-07-30T00:00:00Z',
      '1900-02-29T00:00:00Z',
      '2026-07-30T00:00:00',
      '2026-07-30T00:00:00+14:01',
      null,
      true,
      false,
      0,
      1n,
      new Date(),
      {},
      [],
      Symbol('createdAt'),
    ]) {
      expect(() => PptxDocument.create({ createdAt } as never)).toThrow(TypeError);
    }
  });

  it('creates, edits, clears, rolls back, and reopens presentation modified-at metadata', async () => {
    const readCoreXml = (document: PptxDocument): string => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    const omitted = PptxDocument.create();
    const explicitUndefined = PptxDocument.create({ modifiedAt: undefined } as never);
    const custom = PptxDocument.create({
      modifiedAt: '2024-03-01T01:02:03.456+08:00',
    });

    expect([omitted.modifiedAt, explicitUndefined.modifiedAt, custom.modifiedAt]).toEqual([
      undefined,
      undefined,
      '2024-03-01T01:02:03.456+08:00',
    ]);
    expect(readCoreXml(explicitUndefined)).toBe(readCoreXml(omitted));
    expect(readCoreXml(omitted)).not.toContain('<dcterms:modified');
    expect(readCoreXml(custom)).toMatch(
      /<dcterms:modified\b[^>]*\bxsi:type="dcterms:W3CDTF"[^>]*>2024-03-01T01:02:03\.456\+08:00<\/dcterms:modified>/,
    );
    expect(readCoreXml(custom).match(/xmlns:dcterms=/g)).toHaveLength(1);
    expect(readCoreXml(custom).match(/xmlns:xsi=/g)).toHaveLength(1);

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({
        format,
        modifiedAt: '2026-07-30T01:02:03Z',
      });
      expect(created.modifiedAt).toBe('2026-07-30T01:02:03Z');
      expect(readCoreXml(created)).toMatch(
        /<dcterms:modified\b[^>]*\bxsi:type="dcterms:W3CDTF"[^>]*>2026-07-30T01:02:03Z<\/dcterms:modified>/,
      );
      expect(validatePackage(created.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);
      const reopened = await PptxDocument.open(await created.write());
      expect(reopened.modifiedAt).toBe('2026-07-30T01:02:03Z');
      expect(reopened.format).toBe(format);
    }

    const combined = PptxDocument.create({
      author: 'Combined author',
      company: 'Combined company',
      createdAt: '2026-07-29T00:00:00Z',
      lastModifiedBy: 'Combined editor',
      modifiedAt: '2026-07-30T01:02:03.456+08:00',
      revision: '8',
      subject: 'Combined subject',
      title: 'Combined title',
    });
    expect([
      combined.author,
      combined.company,
      combined.createdAt,
      combined.lastModifiedBy,
      combined.modifiedAt,
      combined.revision,
      combined.subject,
      combined.title,
    ]).toEqual([
      'Combined author',
      'Combined company',
      '2026-07-29T00:00:00Z',
      'Combined editor',
      '2026-07-30T01:02:03.456+08:00',
      '8',
      'Combined subject',
      'Combined title',
    ]);
    expect(readCoreXml(combined)).toMatch(
      /<dcterms:created\b[^>]*\bxsi:type="dcterms:W3CDTF"[^>]*>2026-07-29T00:00:00Z<\/dcterms:created>/,
    );
    expect(readCoreXml(combined)).toMatch(
      /<dcterms:modified\b[^>]*\bxsi:type="dcterms:W3CDTF"[^>]*>2026-07-30T01:02:03\.456\+08:00<\/dcterms:modified>/,
    );
    expect(readCoreXml(combined)).toContain('<dc:creator>Combined author</dc:creator>');
    expect(readCoreXml(combined)).toContain(
      '<cp:lastModifiedBy>Combined editor</cp:lastModifiedBy>',
    );
    expect(readCoreXml(combined)).toContain('<cp:revision>8</cp:revision>');
    expect(readCoreXml(combined)).toContain('<dc:subject>Combined subject</dc:subject>');
    expect(readCoreXml(combined)).toContain('<dc:title>Combined title</dc:title>');
    expect(new TextDecoder().decode(
      combined.opcPackage.requirePart('/docProps/app.xml').bytes,
    )).toContain('<Company>Combined company</Company>');

    const beforeSame = custom.opcPackage.requirePart('/docProps/core.xml').bytes;
    const sameJournal = [...custom.opcPackage.mutations];
    custom.modifiedAt = '2024-03-01T01:02:03.456+08:00';
    expect(custom.opcPackage.requirePart('/docProps/core.xml').bytes).toEqual(beforeSame);
    expect(custom.opcPackage.mutations).toEqual(sameJournal);

    const slide = custom.addSlide();
    const otherParts = new Map(
      custom.opcPackage.parts
        .filter(({ uri }) => uri !== '/docProps/core.xml')
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    custom.createdAt = '2024-02-29T12:34:56Z';
    custom.modifiedAt = '2026-07-30T01:02:03Z';
    expect(custom.modifiedAt).toBe('2026-07-30T01:02:03Z');
    expect(custom.createdAt).toBe('2024-02-29T12:34:56Z');
    expect(custom.author).toBe('@jiayunxie/pptx');
    expect(custom.lastModifiedBy).toBe('@jiayunxie/pptx');
    expect(custom.revision).toBe('1');
    expect(custom.slides[0]).toBe(slide);
    for (const [uri, bytes] of otherParts) {
      expect(custom.opcPackage.requirePart(uri).bytes).toEqual(bytes);
    }

    const firstWrite = await custom.write();
    const firstCore = readCoreXml(custom);
    const secondWrite = await custom.write();
    expect(readCoreXml(custom)).toBe(firstCore);
    expect(await PptxDocument.open(firstWrite).then(({ modifiedAt }) => modifiedAt))
      .toBe('2026-07-30T01:02:03Z');
    expect(await PptxDocument.open(secondWrite).then(({ modifiedAt }) => modifiedAt))
      .toBe('2026-07-30T01:02:03Z');

    const beforeRollback = custom.opcPackage.requirePart('/docProps/core.xml').bytes;
    const rollbackJournal = [...custom.opcPackage.mutations];
    expect(() => custom.transaction(() => {
      custom.modifiedAt = '2025-01-01T00:00:00Z';
      expect(custom.modifiedAt).toBe('2025-01-01T00:00:00Z');
      throw new Error('restore presentation modifiedAt');
    })).toThrow('restore presentation modifiedAt');
    expect(custom.opcPackage.requirePart('/docProps/core.xml').bytes).toEqual(beforeRollback);
    expect(custom.opcPackage.mutations).toEqual(rollbackJournal);
    expect(custom.modifiedAt).toBe('2026-07-30T01:02:03Z');
    expect(custom.createdAt).toBe('2024-02-29T12:34:56Z');
    expect(custom.slides[0]).toBe(slide);

    const reopenedEdited = await PptxDocument.open(await custom.write());
    expect(reopenedEdited.modifiedAt).toBe('2026-07-30T01:02:03Z');
    expect(reopenedEdited.createdAt).toBe('2024-02-29T12:34:56Z');
    reopenedEdited.modifiedAt = undefined;
    expect(reopenedEdited.modifiedAt).toBeUndefined();
    expect(reopenedEdited.createdAt).toBe('2024-02-29T12:34:56Z');
    expect(readCoreXml(reopenedEdited)).not.toContain('<dcterms:modified');
    expect(readCoreXml(reopenedEdited)).toMatch(
      /<dcterms:created\b[^>]*\bxsi:type="dcterms:W3CDTF"[^>]*>2024-02-29T12:34:56Z<\/dcterms:created>/,
    );
    expect(readCoreXml(reopenedEdited)).toContain(
      '<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>',
    );
    expect(readCoreXml(reopenedEdited)).toContain('<cp:revision>1</cp:revision>');
    const reopenedCleared = await PptxDocument.open(await reopenedEdited.write());
    expect(reopenedCleared.modifiedAt).toBeUndefined();
    expect(reopenedCleared.createdAt).toBe('2024-02-29T12:34:56Z');
    expect(reopenedCleared.slides).toHaveLength(1);
  });

  it('rejects malformed presentation modified-at metadata during creation', () => {
    for (const modifiedAt of [
      '',
      ' 2026-07-30T01:02:03Z',
      '1900-02-29T01:02:03Z',
      '2026-07-30T01:02:03',
      '2026-07-30T01:02:03+14:01',
      null,
      true,
      false,
      0,
      1n,
      new Date(),
      {},
      [],
      Symbol('modifiedAt'),
    ]) {
      expect(() => PptxDocument.create({ modifiedAt } as never)).toThrow(TypeError);
    }
    expect(PptxDocument.create().modifiedAt).toBeUndefined();
  });

  it('creates, edits, clears, rolls back, and reopens presentation subject metadata', async () => {
    const readCoreXml = (document: PptxDocument): string => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    const omitted = PptxDocument.create();
    const explicitUndefined = PptxDocument.create({ subject: undefined } as never);
    const custom = PptxDocument.create({ subject: 'Revenue & <Forecast>' });
    const empty = PptxDocument.create({ subject: '' });

    expect([
      omitted.subject,
      explicitUndefined.subject,
      custom.subject,
      empty.subject,
    ]).toEqual([undefined, undefined, 'Revenue & <Forecast>', '']);
    expect(readCoreXml(explicitUndefined)).toBe(readCoreXml(omitted));
    expect(readCoreXml(omitted)).not.toContain('<dc:subject');
    expect(readCoreXml(custom)).toContain(
      '<dc:subject>Revenue &amp; &lt;Forecast&gt;</dc:subject>',
    );
    expect(readCoreXml(empty)).toContain('<dc:subject></dc:subject>');

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({ format, subject: `Subject ${format}` });
      expect(created.subject).toBe(`Subject ${format}`);
      expect(readCoreXml(created)).toContain(
        `<dc:subject>Subject ${format}</dc:subject>`,
      );
      expect(validatePackage(created.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);
      const reopened = await PptxDocument.open(await created.write());
      expect(reopened.subject).toBe(`Subject ${format}`);
      expect(reopened.format).toBe(format);
    }

    const combined = PptxDocument.create({
      author: 'Combined author',
      company: 'Combined company',
      subject: 'Combined subject',
      title: 'Combined title',
    });
    expect([
      combined.author,
      combined.company,
      combined.subject,
      combined.title,
    ]).toEqual([
      'Combined author',
      'Combined company',
      'Combined subject',
      'Combined title',
    ]);
    expect(readCoreXml(combined)).toContain('<dc:creator>Combined author</dc:creator>');
    expect(readCoreXml(combined)).toContain('<dc:title>Combined title</dc:title>');
    expect(readCoreXml(combined)).toContain('<dc:subject>Combined subject</dc:subject>');
    expect(readCoreXml(combined)).toContain(
      '<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>',
    );
    expect(readCoreXml(combined)).toContain('<cp:revision>1</cp:revision>');
    expect(new TextDecoder().decode(
      combined.opcPackage.requirePart('/docProps/app.xml').bytes,
    )).toContain('<Company>Combined company</Company>');

    const beforeSame = custom.opcPackage.requirePart('/docProps/core.xml').bytes;
    const sameJournal = [...custom.opcPackage.mutations];
    custom.subject = 'Revenue & <Forecast>';
    expect(custom.opcPackage.requirePart('/docProps/core.xml').bytes).toEqual(beforeSame);
    expect(custom.opcPackage.mutations).toEqual(sameJournal);

    const slide = custom.addSlide();
    const otherParts = new Map(
      custom.opcPackage.parts
        .filter(({ uri }) => uri !== '/docProps/core.xml')
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    custom.subject = 'Edited subject';
    expect(custom.subject).toBe('Edited subject');
    expect(custom.slides[0]).toBe(slide);
    expect(readCoreXml(custom)).toContain('<dc:subject>Edited subject</dc:subject>');
    expect(readCoreXml(custom)).toContain(
      '<dc:creator>@jiayunxie/pptx</dc:creator>',
    );
    expect(readCoreXml(custom)).toContain(
      '<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>',
    );
    expect(readCoreXml(custom)).toContain('<cp:revision>1</cp:revision>');
    for (const [uri, bytes] of otherParts) {
      expect(custom.opcPackage.requirePart(uri).bytes).toEqual(bytes);
    }

    const beforeRollback = custom.opcPackage.requirePart('/docProps/core.xml').bytes;
    const rollbackJournal = [...custom.opcPackage.mutations];
    expect(() => custom.transaction(() => {
      custom.subject = 'Temporary subject';
      expect(custom.subject).toBe('Temporary subject');
      throw new Error('restore presentation subject');
    })).toThrow('restore presentation subject');
    expect(custom.opcPackage.requirePart('/docProps/core.xml').bytes).toEqual(beforeRollback);
    expect(custom.opcPackage.mutations).toEqual(rollbackJournal);
    expect(custom.subject).toBe('Edited subject');
    expect(custom.slides[0]).toBe(slide);

    const reopenedEdited = await PptxDocument.open(await custom.write());
    expect(reopenedEdited.subject).toBe('Edited subject');
    reopenedEdited.subject = '';
    expect(reopenedEdited.subject).toBe('');
    expect(readCoreXml(reopenedEdited)).toContain('<dc:subject></dc:subject>');
    reopenedEdited.subject = undefined;
    expect(reopenedEdited.subject).toBeUndefined();
    expect(readCoreXml(reopenedEdited)).not.toContain('<dc:subject');
    expect(readCoreXml(reopenedEdited)).toContain(
      '<dc:creator>@jiayunxie/pptx</dc:creator>',
    );
    expect(readCoreXml(reopenedEdited)).toContain(
      '<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>',
    );
    expect(readCoreXml(reopenedEdited)).toContain('<cp:revision>1</cp:revision>');
    const reopenedCleared = await PptxDocument.open(await reopenedEdited.write());
    expect(reopenedCleared.subject).toBeUndefined();
    expect(reopenedCleared.slides).toHaveLength(1);
  });

  it('rejects malformed presentation subject metadata during creation', () => {
    for (const subject of [
      null,
      true,
      false,
      0,
      1,
      {},
      [],
      Symbol('subject'),
      'bad\u0001subject',
    ]) {
      expect(() => PptxDocument.create({ subject: subject as never })).toThrow(TypeError);
    }
  });

  it('creates, edits, clears, rolls back, and reopens presentation revision metadata', async () => {
    const readCoreXml = (document: PptxDocument): string => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/core.xml').bytes,
    );
    const omitted = PptxDocument.create();
    const explicitUndefined = PptxDocument.create({ revision: undefined } as never);
    const custom = PptxDocument.create({ revision: '7' });
    const leading = PptxDocument.create({ revision: '007' });

    expect([
      omitted.revision,
      explicitUndefined.revision,
      custom.revision,
      leading.revision,
    ]).toEqual(['1', '1', '7', '007']);
    expect(readCoreXml(explicitUndefined)).toBe(readCoreXml(omitted));
    expect(readCoreXml(custom)).toContain('<cp:revision>7</cp:revision>');
    expect(readCoreXml(leading)).toContain('<cp:revision>007</cp:revision>');

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({ format, revision: '42' });
      expect(created.revision).toBe('42');
      expect(readCoreXml(created)).toContain('<cp:revision>42</cp:revision>');
      expect(validatePackage(created.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);
      const reopened = await PptxDocument.open(await created.write());
      expect(reopened.revision).toBe('42');
      expect(reopened.format).toBe(format);
    }

    const combined = PptxDocument.create({
      author: 'Combined author',
      company: 'Combined company',
      revision: '8',
      subject: 'Combined subject',
      title: 'Combined title',
    });
    expect([
      combined.author,
      combined.company,
      combined.revision,
      combined.subject,
      combined.title,
    ]).toEqual([
      'Combined author',
      'Combined company',
      '8',
      'Combined subject',
      'Combined title',
    ]);
    expect(readCoreXml(combined)).toContain('<dc:creator>Combined author</dc:creator>');
    expect(readCoreXml(combined)).toContain('<dc:title>Combined title</dc:title>');
    expect(readCoreXml(combined)).toContain('<dc:subject>Combined subject</dc:subject>');
    expect(readCoreXml(combined)).toContain('<cp:revision>8</cp:revision>');
    expect(readCoreXml(combined)).toContain(
      '<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>',
    );
    expect(new TextDecoder().decode(
      combined.opcPackage.requirePart('/docProps/app.xml').bytes,
    )).toContain('<Company>Combined company</Company>');

    const beforeSame = custom.opcPackage.requirePart('/docProps/core.xml').bytes;
    const sameJournal = [...custom.opcPackage.mutations];
    custom.revision = '7';
    expect(custom.opcPackage.requirePart('/docProps/core.xml').bytes).toEqual(beforeSame);
    expect(custom.opcPackage.mutations).toEqual(sameJournal);

    const slide = custom.addSlide();
    const otherParts = new Map(
      custom.opcPackage.parts
        .filter(({ uri }) => uri !== '/docProps/core.xml')
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    custom.revision = '42';
    expect(custom.revision).toBe('42');
    expect(custom.slides[0]).toBe(slide);
    expect(readCoreXml(custom)).toContain('<cp:revision>42</cp:revision>');
    expect(readCoreXml(custom)).toContain(
      '<dc:creator>@jiayunxie/pptx</dc:creator>',
    );
    expect(readCoreXml(custom)).toContain(
      '<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>',
    );
    for (const [uri, bytes] of otherParts) {
      expect(custom.opcPackage.requirePart(uri).bytes).toEqual(bytes);
    }

    custom.revision = '0009';
    expect(custom.revision).toBe('0009');
    expect(readCoreXml(custom)).toContain('<cp:revision>0009</cp:revision>');

    const beforeRollback = custom.opcPackage.requirePart('/docProps/core.xml').bytes;
    const rollbackJournal = [...custom.opcPackage.mutations];
    expect(() => custom.transaction(() => {
      custom.revision = '99';
      expect(custom.revision).toBe('99');
      throw new Error('restore presentation revision');
    })).toThrow('restore presentation revision');
    expect(custom.opcPackage.requirePart('/docProps/core.xml').bytes).toEqual(beforeRollback);
    expect(custom.opcPackage.mutations).toEqual(rollbackJournal);
    expect(custom.revision).toBe('0009');
    expect(custom.slides[0]).toBe(slide);

    const reopenedEdited = await PptxDocument.open(await custom.write());
    expect(reopenedEdited.revision).toBe('0009');
    reopenedEdited.revision = undefined;
    expect(reopenedEdited.revision).toBeUndefined();
    expect(readCoreXml(reopenedEdited)).not.toContain('<cp:revision');
    expect(readCoreXml(reopenedEdited)).toContain(
      '<dc:creator>@jiayunxie/pptx</dc:creator>',
    );
    expect(readCoreXml(reopenedEdited)).toContain(
      '<cp:lastModifiedBy>@jiayunxie/pptx</cp:lastModifiedBy>',
    );
    const reopenedCleared = await PptxDocument.open(await reopenedEdited.write());
    expect(reopenedCleared.revision).toBeUndefined();
    expect(reopenedCleared.slides).toHaveLength(1);
  });

  it('rejects malformed presentation revision metadata during creation', () => {
    for (const revision of [
      '',
      ' ',
      '+1',
      '-1',
      '1.0',
      '1e3',
      '１２',
      null,
      true,
      false,
      0,
      1,
      1n,
      {},
      [],
      Symbol('revision'),
    ]) {
      expect(() => PptxDocument.create({ revision: revision as never })).toThrow(TypeError);
    }
  });

  it('creates, edits, clears, rolls back, and reopens presentation company metadata', async () => {
    const readAppXml = (document: PptxDocument): string => new TextDecoder().decode(
      document.opcPackage.requirePart('/docProps/app.xml').bytes,
    );
    const omitted = PptxDocument.create();
    const explicitUndefined = PptxDocument.create({ company: undefined } as never);
    const custom = PptxDocument.create({ company: 'Acme & <Partners>' });
    const empty = PptxDocument.create({ company: '' });

    expect([
      omitted.company,
      explicitUndefined.company,
      custom.company,
      empty.company,
    ]).toEqual([undefined, undefined, 'Acme & <Partners>', '']);
    expect(readAppXml(explicitUndefined)).toBe(readAppXml(omitted));
    expect(readAppXml(omitted)).not.toContain('<Company');
    expect(readAppXml(custom)).toContain(
      '<Company>Acme &amp; &lt;Partners&gt;</Company><AppVersion>1.0</AppVersion>',
    );
    expect(readAppXml(empty)).toContain(
      '<Company></Company><AppVersion>1.0</AppVersion>',
    );

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({ format, company: `Company ${format}` });
      expect(created.company).toBe(`Company ${format}`);
      expect(readAppXml(created)).toContain(`<Company>Company ${format}</Company>`);
      expect(validatePackage(created.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);
      const reopened = await PptxDocument.open(await created.write());
      expect(reopened.company).toBe(`Company ${format}`);
      expect(reopened.format).toBe(format);
    }

    const combined = PptxDocument.create({
      author: 'Combined author',
      company: 'Combined company',
      rtlMode: true,
      title: 'Combined title',
    });
    expect([
      combined.author,
      combined.company,
      combined.rtlMode,
      combined.title,
    ]).toEqual(['Combined author', 'Combined company', true, 'Combined title']);
    expect(readAppXml(combined)).toContain('<Company>Combined company</Company>');
    expect(readAppXml(combined)).toContain('<Application>@jiayunxie/pptx</Application>');
    expect(readAppXml(combined)).toContain('<PresentationFormat>Custom</PresentationFormat>');
    expect(readAppXml(combined)).toContain('<AppVersion>1.0</AppVersion>');

    const beforeSame = custom.opcPackage.requirePart('/docProps/app.xml').bytes;
    const sameJournal = [...custom.opcPackage.mutations];
    custom.company = 'Acme & <Partners>';
    expect(custom.opcPackage.requirePart('/docProps/app.xml').bytes).toEqual(beforeSame);
    expect(custom.opcPackage.mutations).toEqual(sameJournal);

    const slide = custom.addSlide();
    const otherParts = new Map(
      custom.opcPackage.parts
        .filter(({ uri }) => uri !== '/docProps/app.xml')
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    custom.company = 'Edited company';
    expect(custom.company).toBe('Edited company');
    expect(custom.slides[0]).toBe(slide);
    expect(readAppXml(custom)).toContain('<Company>Edited company</Company>');
    expect(readAppXml(custom)).toContain('<Application>@jiayunxie/pptx</Application>');
    expect(readAppXml(custom)).toContain('<PresentationFormat>Custom</PresentationFormat>');
    expect(readAppXml(custom)).toContain('<AppVersion>1.0</AppVersion>');
    for (const [uri, bytes] of otherParts) {
      expect(custom.opcPackage.requirePart(uri).bytes).toEqual(bytes);
    }

    const beforeRollback = custom.opcPackage.requirePart('/docProps/app.xml').bytes;
    const rollbackJournal = [...custom.opcPackage.mutations];
    expect(() => custom.transaction(() => {
      custom.company = 'Temporary company';
      expect(custom.company).toBe('Temporary company');
      throw new Error('restore presentation company');
    })).toThrow('restore presentation company');
    expect(custom.opcPackage.requirePart('/docProps/app.xml').bytes).toEqual(beforeRollback);
    expect(custom.opcPackage.mutations).toEqual(rollbackJournal);
    expect(custom.company).toBe('Edited company');
    expect(custom.slides[0]).toBe(slide);

    const reopenedEdited = await PptxDocument.open(await custom.write());
    expect(reopenedEdited.company).toBe('Edited company');
    reopenedEdited.company = '';
    expect(reopenedEdited.company).toBe('');
    reopenedEdited.company = undefined;
    expect(reopenedEdited.company).toBeUndefined();
    expect(readAppXml(reopenedEdited)).not.toContain('<Company');
    expect(readAppXml(reopenedEdited)).toContain(
      '<Application>@jiayunxie/pptx</Application>',
    );
    expect(readAppXml(reopenedEdited)).toContain('<AppVersion>1.0</AppVersion>');
    const reopenedCleared = await PptxDocument.open(await reopenedEdited.write());
    expect(reopenedCleared.company).toBeUndefined();
    expect(reopenedCleared.slides).toHaveLength(1);
  });

  it('rejects malformed presentation company metadata during creation', () => {
    for (const company of [
      null,
      true,
      false,
      0,
      1,
      {},
      [],
      Symbol('company'),
      'bad\u0001company',
    ]) {
      expect(() => PptxDocument.create({ company: company as never })).toThrow(TypeError);
    }
  });

  it('creates, edits, clears, rolls back, and reopens presentation RTL independently', async () => {
    const readPresentationXml = (document: PptxDocument): string => new TextDecoder().decode(
      document.opcPackage.requirePart(document.presentationPartUri).bytes,
    );
    const omitted = PptxDocument.create();
    const enabled = PptxDocument.create({ rtlMode: true });
    const disabled = PptxDocument.create({ rtlMode: false });
    expect([omitted.rtlMode, enabled.rtlMode, disabled.rtlMode]).toEqual([undefined, true, false]);
    expect(readPresentationXml(enabled)).toMatch(/<p:presentation[^>]* rtl="1"/);
    expect(readPresentationXml(disabled)).toMatch(/<p:presentation[^>]* rtl="0"/);
    expect(readPresentationXml(omitted)).not.toMatch(/<p:presentation[^>]*\srtl=/);

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({ format, rtlMode: true });
      expect(created.rtlMode).toBe(true);
      expect(validatePackage(created.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
    }

    const slide = enabled.addSlide();
    const shape = slide.addRichText([
      { rtl: true, runs: [{ text: 'Paragraph RTL' }] },
      { rtl: false, runs: [{ text: 'Paragraph LTR' }] },
    ]);
    enabled.duplicateSlide(0);
    const paragraphState = shape.richText.map(({ rtl }) => rtl);

    enabled.rtlMode = false;
    expect(enabled.rtlMode).toBe(false);
    expect(readPresentationXml(enabled)).toMatch(/<p:presentation[^>]* rtl="0"/);
    expect(shape.richText.map(({ rtl }) => rtl)).toEqual(paragraphState);

    const beforeRollback = enabled.opcPackage.requirePart(enabled.presentationPartUri).bytes;
    const journal = [...enabled.opcPackage.mutations];
    expect(() =>
      enabled.transaction(() => {
        enabled.rtlMode = true;
        throw new Error('restore presentation RTL');
      }),
    ).toThrow('restore presentation RTL');
    expect(enabled.opcPackage.requirePart(enabled.presentationPartUri).bytes).toEqual(beforeRollback);
    expect(enabled.opcPackage.mutations).toEqual(journal);
    expect(enabled.slides[0]).toBe(slide);
    expect(enabled.slides[0]!.shapes[0]).toBe(shape);
    expect(enabled.rtlMode).toBe(false);

    const reopenedTrue = await PptxDocument.open(await PptxDocument.create({ rtlMode: true }).write());
    const reopenedFalse = await PptxDocument.open(await disabled.write());
    expect([reopenedTrue.rtlMode, reopenedFalse.rtlMode]).toEqual([true, false]);

    enabled.rtlMode = undefined;
    expect(enabled.rtlMode).toBeUndefined();
    expect(readPresentationXml(enabled)).not.toMatch(/<p:presentation[^>]*\srtl=/);
    expect(shape.richText.map(({ rtl }) => rtl)).toEqual(paragraphState);
    const reopened = await PptxDocument.open(await enabled.write());
    expect(reopened.rtlMode).toBeUndefined();
    expect(reopened.slides).toHaveLength(2);
    expect((reopened.slides[0]!.shapes[0] as ShapeModel).richText.map(({ rtl }) => rtl))
      .toEqual(paragraphState);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('rejects malformed presentation RTL before returning a created document', () => {
    for (const rtlMode of [null, 0, 'true', {}, [], Symbol('rtl')]) {
      expect(() => PptxDocument.create({ rtlMode: rtlMode as never })).toThrow(TypeError);
    }
  });

  it('accepts custom slide size boundaries and rejects malformed or out-of-range dimensions', () => {
    expect(() =>
      PptxDocument.create({ slideSize: Object.freeze({ width: inches(1), height: inches(56) }) }),
    ).not.toThrow();

    for (const slideSize of [
      null,
      [],
      {},
      { width: Number.NaN, height: inches(1) },
      { width: inches(1), height: Number.POSITIVE_INFINITY },
    ]) {
      expect(() => PptxDocument.create({ slideSize: slideSize as never })).toThrow(TypeError);
    }
    for (const slideSize of [
      { width: inches(0.99), height: inches(1) },
      { width: inches(1), height: inches(56.01) },
    ]) {
      expect(() => PptxDocument.create({ slideSize })).toThrow(RangeError);
    }
  });

  it('rejects unknown create options before returning a document', () => {
    expect(() => PptxDocument.create({ format: 'pdf' as PresentationFormat })).toThrow(
      /Unsupported presentation format/,
    );
    expect(() => PptxDocument.create({ slideSize: 'a4' as '16:9' })).toThrow(
      /Unsupported built-in slide size/,
    );
  });

  it('creates, edits, and round-trips a basic text shape with stable identity', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const text = slide.addText(' Hello & <world> ', {
      name: 'Heading "A"',
      x: inches(1),
      y: inches(1.5),
      width: inches(4),
      height: inches(1),
      rotation: degrees(15),
      flipHorizontal: true,
    });

    expect(text).toBeInstanceOf(ShapeModel);
    expect(text.kind).toBe('text');
    expect(text.text).toBe(' Hello & <world> ');
    expect(text.name).toBe('Heading "A"');
    expect(text.transform).toMatchObject({
      x: inches(1),
      y: inches(1.5),
      width: inches(4),
      height: inches(1),
      rotation: degrees(15),
      flipHorizontal: true,
      flipVertical: false,
    });
    expect(slide.shapes[0]).toBe(text);

    text.text = 'Updated & safe';
    text.setTransform({ y: inches(2), flipVertical: true });
    expect(slide.shapes[0]).toBe(text);

    const reopened = await PptxDocument.open(await document.write());
    const roundTripped = reopened.slides[0]!.shapes[0] as ShapeModel;
    expect(roundTripped.text).toBe('Updated & safe');
    expect(roundTripped.name).toBe('Heading "A"');
    expect(roundTripped.transform).toMatchObject({
      x: inches(1),
      y: inches(2),
      width: inches(4),
      height: inches(1),
      rotation: degrees(15),
      flipHorizontal: true,
      flipVertical: true,
    });
  });

  it('creates, edits, duplicates, rolls back, and reopens table cell text fit', async () => {
    const document = PptxDocument.create({ slideSize: 'wide' });
    const slide = document.addSlide();
    const table = slide.addTable([[
      'String',
      { text: 'Omitted' },
      {
        text: 'Undefined',
        options: { fit: undefined } as unknown as AddTableCellOptions,
      },
      { text: 'None', options: { fit: 'none' } },
      { text: 'Shrink', options: { fit: 'shrink' } },
      { text: 'Resize', options: {
        fit: 'resize',
        textDirection: 'vert270',
        valign: 'middle',
        fill: { kind: 'solid', color: { kind: 'srgb', value: 'FFF2CC' } },
      } },
    ]], {
      columnWidths: inches(1.5),
      rowHeights: inches(1),
    });

    expect(table.rows[0]!.cells.map(({ textFit }) => textFit)).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
      'shrink',
      'resize',
    ]);
    expect(table.rows[0]!.cells[5]!.textDirection).toBe('vert270');
    expect(table.rows[0]!.cells[5]!.verticalAlignment).toBe('middle');

    const sourceBytes = document.opcPackage.requirePart(slide.partUri).bytes;
    const sourceXml = new TextDecoder().decode(sourceBytes);
    expect(sourceXml.match(/<a:normAutofit\/>/g)).toHaveLength(1);
    expect(sourceXml.match(/<a:spAutoFit\/>/g)).toHaveLength(1);
    expect(sourceXml).not.toContain('<a:noAutofit/>');

    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;
    expect(duplicateTable.rows[0]!.cells.map(({ textFit }) => textFit))
      .toEqual([undefined, undefined, undefined, undefined, 'shrink', 'resize']);

    table.setCellText(0, 4, 'Shrink edited');
    table.setCellTextDirection(0, 5, 'wordArtVert');
    table.setCellTextFit(0, 4, 'none');
    table.setCellTextFit(0, 3, 'resize');
    expect(table.rows[0]!.cells.map(({ textFit }) => textFit)).toEqual([
      undefined,
      undefined,
      undefined,
      'resize',
      undefined,
      'resize',
    ]);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    expect(() => document.transaction(() => {
      table.setCellTextFit(0, 3, 'shrink');
      table.setCellTextFit(0, 5, undefined);
      throw new Error('restore created fits');
    })).toThrow('restore created fits');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(journal);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    expect(reopenedTable.rows[0]!.cells.map(({ textFit }) => textFit)).toEqual([
      undefined,
      undefined,
      undefined,
      'resize',
      undefined,
      'resize',
    ]);
    expect(reopenedTable.rows[0]!.cells[4]!.text).toBe('Shrink edited');
    expect(reopenedTable.rows[0]!.cells[5]!.textDirection).toBe('wordArtVert');
    expect((reopened.slides[1]!.shapes[0] as TableModel).rows[0]!.cells[4]!.textFit)
      .toBe('shrink');
  });

  it('creates and edits table-cell hyperlinks through the public SDK surface', async () => {
    const document = PptxDocument.create();
    const source = document.addSlide();
    const target = document.addSlide();
    const url: Hyperlink = {
      url: 'https://example.com?a=1&b=2',
      tooltip: 'Visit & learn',
    };
    const options: AddTableCellOptions = { hyperlink: url };
    const table = source.addTable([[
      { text: 'URL', options },
      { text: 'Slide', options: { hyperlink: { slide: 2, tooltip: '' } } },
      'Plain',
    ]], { name: 'SDK cell hyperlinks' });
    const cell: TableCell = table.rows[0]!.cells[0]!;
    const hyperlink: Hyperlink | undefined = cell.hyperlink;

    expect(hyperlink).toEqual(url);
    expect(table.rows[0]!.cells.map((candidate) => candidate.hyperlink)).toEqual([
      url,
      { slide: 2, tooltip: '' },
      undefined,
    ]);
    expect(source.relationships.filter(({ type }) => type.endsWith('/hyperlink')))
      .toEqual([expect.objectContaining({
        target: 'https://example.com?a=1&b=2',
        targetMode: 'External',
      })]);
    expect(source.relationships.find(({ resolvedTarget }) => resolvedTarget === target.partUri))
      .toBeDefined();
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error'))
      .toEqual([]);

    const beforeInvalid = document.opcPackage.requirePart(source.partUri).bytes;
    const relationships = source.relationships;
    const journal = [...document.opcPackage.mutations];
    expect(() => source.addTable([[
      { text: 'First', options: { hyperlink: { url: 'https://first.example' } } },
      { text: 'Invalid', options: { hyperlink: { slide: 99 } } },
    ]])).toThrow('Table cell 0,1 hyperlink slide 99 is out of range');
    expect(document.opcPackage.requirePart(source.partUri).bytes).toEqual(beforeInvalid);
    expect(source.relationships).toEqual(relationships);
    expect(document.opcPackage.mutations).toEqual(journal);

    const beforeNoOp = document.opcPackage.requirePart(source.partUri).bytes;
    const beforeNoOpJournal = [...document.opcPackage.mutations];
    table.setCellHyperlink(0, 0, url);
    expect(document.opcPackage.requirePart(source.partUri).bytes).toEqual(beforeNoOp);
    expect(document.opcPackage.mutations).toEqual(beforeNoOpJournal);

    const editedUrl: Hyperlink = {
      url: 'https://sdk-edit.example?a=1&b=2',
      tooltip: '',
    };
    table.setCellHyperlink(0, 0, editedUrl);
    table.setCellHyperlink(0, 1, { slide: 2, tooltip: 'SDK target' });
    table.setCellHyperlink(0, 2, { url: 'https://temporary.example' });
    expect(table.rows[0]!.cells.map((candidate) => candidate.hyperlink)).toEqual([
      editedUrl,
      { slide: 2, tooltip: 'SDK target' },
      { url: 'https://temporary.example' },
    ]);
    table.setCellHyperlink(0, 2, undefined);
    expect(source.shapes[0]).toBe(table);

    table.setCellText(0, 0, 'URL edited');
    document.moveSlide(document.slides.indexOf(target), 0);
    expect(table.rows[0]!.cells[1]!.hyperlink).toEqual({
      slide: 1,
      tooltip: 'SDK target',
    });
    document.moveSlide(0, document.slides.length - 1);
    const duplicate = document.duplicateSlide(document.slides.indexOf(source));
    expect((duplicate.shapes[0] as TableModel).rows[0]!.cells[1]!.hyperlink)
      .toEqual({
        slide: document.slides.indexOf(target) + 1,
        tooltip: 'SDK target',
      });

    const reopened = await PptxDocument.open(await document.write());
    const reopenedSource = reopened.slides.find(({ partUri }) => partUri === source.partUri)!;
    const reopenedTable = reopenedSource.shapes[0] as TableModel;
    expect(reopenedTable.rows[0]!.cells.map((candidate) => candidate.hyperlink)).toEqual([
      editedUrl,
      { slide: reopened.slides.findIndex(({ partUri }) => partUri === target.partUri) + 1,
        tooltip: 'SDK target' },
      undefined,
    ]);

    if (false) {
      const invalid: readonly AddTableCellOptions[] = [
        {
          // @ts-expect-error table-cell hyperlink requires exactly one target
          hyperlink: {},
        },
        {
          // @ts-expect-error table-cell hyperlink branches are mutually exclusive
          hyperlink: { url: 'https://example.com', slide: 2 },
        },
        {
          // @ts-expect-error table-cell hyperlink URL must be a string
          hyperlink: { url: 42 },
        },
        {
          // @ts-expect-error table-cell hyperlink slide must be numeric
          hyperlink: { slide: '2' },
        },
        {
          // @ts-expect-error table-cell hyperlink tooltip must be a string
          hyperlink: { slide: 2, tooltip: 7 },
        },
        {
          hyperlink: {
            url: 'https://example.com',
            // @ts-expect-error relationship IDs are internal
            _rId: 'rId9',
          },
        },
      ];
      // @ts-expect-error there is no table-level hyperlink default
      source.addTable([['A']], { hyperlink: url });
      // @ts-expect-error table-cell hyperlink editor requires a supported hyperlink
      table.setCellHyperlink(0, 0, {});
      // @ts-expect-error table-cell hyperlink editor URL and slide targets are exclusive
      table.setCellHyperlink(0, 0, { url: 'https://example.com', slide: 2 });
      // @ts-expect-error table-cell hyperlink snapshots are readonly
      cell.hyperlink = { url: 'https://mutated.example' };
      expect(invalid).toHaveLength(6);
    }
  });

  it('creates and edits rich table-cell text through the public SDK surface', async () => {
    const document = PptxDocument.create();
    const source = document.addSlide();
    document.addSlide();
    const paragraphs: readonly RichTextParagraph[] = [{
      align: 'center',
      runs: [
        { text: 'SDK rich', style: { bold: true } },
        { text: ' link', style: { hyperlink: { url: 'https://sdk-rich.example' } } },
      ],
    }, {
      runs: [{ text: 'Second', style: { italic: true } }],
    }];
    const richCell: AddTableCell = { text: paragraphs };
    const rows: readonly (readonly AddTableCellInput[])[] = [[
      richCell,
      'one\r\ntwo\n',
    ]];
    const table = source.addTable(rows, { name: 'SDK rich cells' });
    const cell: TableCell = table.rows[0]!.cells[0]!;
    const snapshot: readonly RichTextParagraph[] = cell.richText;

    expect(cell.text).toBe('SDK rich link\nSecond');
    expect(snapshot.map(({ runs }) => runs.map(({ text }) => text)))
      .toEqual([['SDK rich', ' link'], ['Second']]);
    expect(table.rows[0]!.cells[1]!.text).toBe('one\ntwo\n');
    expect(table.rows[0]!.cells[1]!.richText).toHaveLength(3);

    table.setCellRichText(0, 0, [{ runs: [
      { text: 'Edited URL', style: { hyperlink: {
        url: 'https://sdk-edited.example',
        tooltip: '',
      } } },
      { text: ' target', style: { hyperlink: { slide: 2, tooltip: 'Target' } } },
    ] }]);
    expect(table.rows[0]!.cells[0]!.text).toBe('Edited URL target');
    expect(table.rows[0]!.cells[0]!.richText[0]!.runs.map(
      ({ style }) => style?.hyperlink,
    )).toEqual([
      { url: 'https://sdk-edited.example', tooltip: '' },
      { slide: 2, tooltip: 'Target' },
    ]);
    expect(validatePackage(document.opcPackage).filter(
      ({ severity }) => severity === 'error',
    )).toEqual([]);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    )!;
    expect(reopenedTable.rows[0]!.cells[0]!.text).toBe('Edited URL target');
    expect(reopenedTable.rows[0]!.cells[0]!.richText[0]!.runs[1]!.style?.hyperlink)
      .toEqual({ slide: 2, tooltip: 'Target' });

    if (false) {
      // @ts-expect-error a bare rich-text array is not a table cell
      const bareCell: AddTableCellInput = [{ runs: [{ text: 'Bare' }] }];
      // @ts-expect-error table-cell richText is a readonly snapshot property
      cell.richText = [];
      // @ts-expect-error rich replacement requires paragraph arrays
      table.setCellRichText(0, 0, 'Plain');
      const recursive: AddTableCell = {
        // @ts-expect-error PptxGenJS recursive table-cell objects are not native paragraphs
        text: [{ text: 'Nested', options: {} }],
      };
      void bareCell;
      void recursive;
    }
  });

  it('creates table-cell text style defaults through the public SDK and root surface', async () => {
    const tableOptions: AddTableOptions = {
      name: 'SDK table-cell text style defaults',
      fontFamily: 'Aptos',
      fontSize: 18.25,
      bold: true,
      color: { kind: 'scheme', value: 'accent1' },
      spacing: {
        before: 6,
        after: 8,
        line: { kind: 'multiple', factor: 1.5 },
      },
    };
    const cellOptions: AddTableCellOptions = {
      fontFamily: 'Courier New',
      fontSize: 10,
      bold: false,
      color: { kind: 'srgb', value: '00AA00' },
      spacing: { before: 3 },
    };
    const document = PptxDocument.create();
    const table = document.addSlide().addTable([[
      'Inherited',
      { text: 'Cell override', options: cellOptions },
    ]], tableOptions);

    expect(table.rows[0]!.cells[0]!.richText[0]!.runs[0]!.style).toMatchObject({
      fontFamily: 'Aptos',
      fontSize: 18.25,
      bold: true,
      color: { kind: 'scheme', value: 'accent1' },
    });
    expect(table.rows[0]!.cells[1]!.richText[0]).toMatchObject({
      spacing: {
        before: 3,
        after: 8,
        line: { kind: 'multiple', factor: 1.5 },
      },
      runs: [{
        text: 'Cell override',
        style: {
          fontFamily: 'Courier New',
          fontSize: 10,
          bold: false,
          color: { kind: 'srgb', value: '00AA00' },
        },
      }],
    });
    expect(validatePackage(document.opcPackage).filter(
      ({ severity }) => severity === 'error',
    )).toEqual([]);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    expect(reopenedTable.rows[0]!.cells.map(({ richText }) => richText))
      .toEqual(table.rows[0]!.cells.map(({ richText }) => richText));

    if (false) {
      const tableAlias: AddTableOptions = {
        // @ts-expect-error native table options use fontFamily, not fontFace
        fontFace: 'Aptos',
      };
      const tableBold: AddTableOptions = {
        // @ts-expect-error table bold must be boolean
        bold: 1,
      };
      const cellSpacingAlias: AddTableCellOptions = {
        // @ts-expect-error native spacing uses structured ParagraphSpacing
        paraSpaceAfter: 6,
      };
      const cellBold: AddTableCellOptions = {
        // @ts-expect-error table-cell bold must be boolean
        bold: 'true',
      };
      void [tableAlias, tableBold, cellSpacingAlias, cellBold];
    }
  });

  it('creates, edits, duplicates, rolls back, and reopens a basic table', async () => {
    const document = PptxDocument.create({ slideSize: 'wide' });
    const slide = document.addSlide();
    const otherParts = new Map(
      document.opcPackage.parts
        .filter(({ uri }) => uri !== slide.partUri)
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    const sourceColor = { kind: 'srgb' as const, value: '#D9EAF7' };
    const sourceFill = {
      kind: 'solid' as const,
      color: sourceColor,
      transparency: 33.3334,
    };
    const sourceBorderColor = { kind: 'srgb' as const, value: '#C00000' };
    const sourceBorder = {
      kind: 'line' as const,
      color: sourceBorderColor,
      width: 2,
      style: 'solid' as const,
    };
    const sourceMargin = { top: 4, left: 8 };
    const sourceOptions: AddTableCellOptions = {
      border: sourceBorder,
      fill: sourceFill,
      margin: sourceMargin,
      valign: 'top',
    };
    const sourceCell = { text: 'Region', options: sourceOptions };
    const rows: readonly (readonly AddTableCellInput[])[] = [
      [
        sourceCell,
        { text: 'Revenue', options: {
          border: {
            kind: 'line',
            color: { kind: 'scheme', value: 'accent2' },
            width: 2,
            style: 'dash',
          },
          fill: {
            kind: 'solid',
            color: { kind: 'scheme', value: 'accent2' },
            transparency: 25,
          },
          margin: [1, 2, 3, 4],
          valign: 'middle',
        } },
        { text: 'Growth', options: {
          border: { kind: 'none' },
          fill: { kind: 'none' },
          margin: 0,
          valign: 'bottom',
        } },
      ],
      [{ text: 'East', options: { margin: { right: -2 } } }, { text: '$1.2M', options: {
        border: {
          kind: 'line',
          color: { kind: 'srgb', value: '0000FF' },
          width: 0,
          style: 'solid',
        },
        fill: {
          kind: 'solid',
          color: { kind: 'srgb', value: '445566' },
          transparency: 100,
        },
        margin: {},
      } }, { text: '12%', options: { border: [
        {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent1' },
          width: 1.5,
          style: 'dash',
        },
        undefined,
        {
          kind: 'line',
          color: { kind: 'srgb', value: '00FF00' },
          width: 0,
        },
        { kind: 'none' },
      ] } }],
      [{ text: 'West', options: { border: {
        top: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent3' },
          width: 1,
          style: 'dash',
        },
        left: { kind: 'none' },
      } } }, '$980K', { text: '' }],
    ];
    const tableOptions: AddTableOptions = {
      name: 'Revenue table',
      x: inches(1),
      y: inches(1.25),
      width: inches(8),
      height: inches(2.25),
      columnWidths: [inches(2), inches(4), inches(2)],
      rowHeights: [inches(0.5), inches(0.75), inches(1)],
      valign: 'middle',
    };
    const table = slide.addTable(rows, tableOptions);

    expect(table).toBeInstanceOf(TableModel);
    expect(table.name).toBe('Revenue table');
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows.map(({ cells }) => cells.map(({ text }) => text))).toEqual([
      ['Region', 'Revenue', 'Growth'],
      ['East', '$1.2M', '12%'],
      ['West', '$980K', ''],
    ]);
    sourceCell.text = 'MUTATED';
    sourceColor.value = '000000';
    sourceFill.transparency = 1;
    sourceBorderColor.value = '000000';
    sourceBorder.width = 9;
    sourceMargin.top = 99;
    sourceMargin.left = 99;
    expect(table.rows[0]!.cells[0]!.text).toBe('Region');
    expect(table.rows.map(({ cells }) => cells.map(({ margins }) => margins))).toEqual([
      [
        { top: 4, right: 7.2, bottom: 3.6, left: 8 },
        { top: 1, right: 2, bottom: 3, left: 4 },
        { top: 0, right: 0, bottom: 0, left: 0 },
      ],
      [
        { top: 3.6, right: -2, bottom: 3.6, left: 7.2 },
        { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 },
        { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 },
      ],
      [
        { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 },
        { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 },
        { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 },
      ],
    ]);
    expect(table.rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment))).toEqual([
      ['top', 'middle', 'bottom'],
      ['middle', 'middle', 'middle'],
      ['middle', 'middle', 'middle'],
    ]);
    expect(table.rows.map(({ cells }) => cells.map(({ fill }) => fill))).toEqual([
      [
        {
          kind: 'solid',
          color: { kind: 'srgb', value: 'D9EAF7' },
          transparency: 33.333,
        },
        {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 25,
        },
        { kind: 'none' },
      ],
      [
        undefined,
        {
          kind: 'solid',
          color: { kind: 'srgb', value: '445566' },
          transparency: 100,
        },
        undefined,
      ],
      [undefined, undefined, undefined],
    ]);
    const borderLine = (
      color: { kind: 'srgb' | 'scheme'; value: string },
      width: number,
      style?: 'solid' | 'dash',
    ) => ({ kind: 'line' as const, color, width, ...(style ? { style } : {}) });
    const noBorders = () => ({
      top: { kind: 'none' as const },
      right: { kind: 'none' as const },
      bottom: { kind: 'none' as const },
      left: { kind: 'none' as const },
    });
    expect(table.rows[0]!.cells[0]!.borders).toEqual({
      top: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
      right: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
      bottom: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
      left: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
    });
    expect(table.rows[0]!.cells[1]!.borders).toEqual({
      top: borderLine({ kind: 'scheme', value: 'accent2' }, 2, 'dash'),
      right: borderLine({ kind: 'scheme', value: 'accent2' }, 2, 'dash'),
      bottom: borderLine({ kind: 'scheme', value: 'accent2' }, 2, 'dash'),
      left: borderLine({ kind: 'scheme', value: 'accent2' }, 2, 'dash'),
    });
    expect(table.rows[0]!.cells[2]!.borders).toEqual(noBorders());
    expect(table.rows[1]!.cells[1]!.borders).toEqual({
      top: borderLine({ kind: 'srgb', value: '0000FF' }, 0, 'solid'),
      right: borderLine({ kind: 'srgb', value: '0000FF' }, 0, 'solid'),
      bottom: borderLine({ kind: 'srgb', value: '0000FF' }, 0, 'solid'),
      left: borderLine({ kind: 'srgb', value: '0000FF' }, 0, 'solid'),
    });
    expect(table.rows[1]!.cells[2]!.borders).toEqual({
      top: borderLine({ kind: 'scheme', value: 'accent1' }, 1.5, 'dash'),
      right: { kind: 'none' },
      bottom: borderLine({ kind: 'srgb', value: '00FF00' }, 0),
      left: { kind: 'none' },
    });
    expect(table.rows[2]!.cells[0]!.borders).toEqual({
      top: borderLine({ kind: 'scheme', value: 'accent3' }, 1, 'dash'),
      right: { kind: 'none' },
      bottom: { kind: 'none' },
      left: { kind: 'none' },
    });
    expect(table.transform).toMatchObject({
      x: inches(1),
      y: inches(1.25),
      width: inches(8),
      height: inches(2.25),
    });
    expect(table.columnWidths).toEqual([
      inches(2),
      inches(4),
      inches(2),
    ]);
    expect(table.rowHeights).toEqual([
      inches(0.5),
      inches(0.75),
      inches(1),
    ]);
    expect(table.rows[0]!.cells[0]).toMatchObject({
      text: 'Region',
      margins: { top: 4, right: 7.2, bottom: 3.6, left: 8 },
      borders: {
        top: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
        right: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
        bottom: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
        left: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
      },
    });
    expect(table.rows[0]!.cells[0]!.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'D9EAF7' },
      transparency: 33.333,
    });
    expect(table.rows[0]!.cells[0]!.textDirection).toBeUndefined();
    expect(table.rows[0]!.cells[0]!.textFit).toBeUndefined();
    expect(table.rows[0]!.cells[0]!.verticalAlignment).toBe('top');
    const createdTableXml = new TextDecoder().decode(
      document.opcPackage.requirePart(slide.partUri).bytes,
    );
    expect([...createdTableXml.matchAll(/<a:gridCol w="(\d+)"\/>/g)]
      .map((match) => Number(match[1]))).toEqual([
      inches(2),
      inches(4),
      inches(2),
    ]);
    expect([...createdTableXml.matchAll(/<a:tr h="(\d+)">/g)]
      .map((match) => Number(match[1]))).toEqual([
      inches(0.5),
      inches(0.75),
      inches(1),
    ]);
    expect(createdTableXml).toContain('<a:ext cx="7315200" cy="2057400"/>');
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
    for (const [uri, bytes] of otherParts) {
      expect(document.opcPackage.requirePart(uri).bytes).toEqual(bytes);
    }

    table.setCellMargins(0, 0, { bottom: 9 });
    expect(table.rows[0]!.cells[0]!.margins).toEqual({ bottom: 9 });
    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;
    const originalRows = duplicateTable.rows;
    expect(duplicateTable.rows[0]!.cells[0]!.margins).toEqual({ bottom: 9 });
    const nonTargetHashes = new Map(
      document.opcPackage.parts
        .filter(({ uri }) => uri !== slide.partUri)
        .map(({ uri, bytes }) => [uri, hash(bytes)]),
    );
    const duplicateTableXml = new TextDecoder().decode(
      document.opcPackage.requirePart(duplicate.partUri).bytes,
    );
    expect([...duplicateTableXml.matchAll(/<a:gridCol w="(\d+)"\/>/g)]
      .map((match) => Number(match[1]))).toEqual([
      inches(2),
      inches(4),
      inches(2),
    ]);
    expect([...duplicateTableXml.matchAll(/<a:tr h="(\d+)">/g)]
      .map((match) => Number(match[1]))).toEqual([
      inches(0.5),
      inches(0.75),
      inches(1),
    ]);
    expect(duplicateTable.columnWidths).toEqual([
      inches(2),
      inches(4),
      inches(2),
    ]);
    expect(duplicateTable.rowHeights).toEqual([
      inches(0.5),
      inches(0.75),
      inches(1),
    ]);
    expect(duplicateTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual([
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'D9EAF7' },
        transparency: 33.333,
      },
      {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 25,
      },
      { kind: 'none' },
    ]);
    expect(duplicateTable.rows[0]!.cells[0]!.borders).toEqual({
      top: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
      right: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
      bottom: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
      left: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
    });
    expect(duplicateTable.rows[1]!.cells[2]!.borders).toEqual({
      top: borderLine({ kind: 'scheme', value: 'accent1' }, 1.5, 'dash'),
      right: { kind: 'none' },
      bottom: borderLine({ kind: 'srgb', value: '00FF00' }, 0),
      left: { kind: 'none' },
    });
    expect(duplicateTable.rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment))).toEqual([
      ['top', 'middle', 'bottom'],
      ['middle', 'middle', 'middle'],
      ['middle', 'middle', 'middle'],
    ]);

    table.setCellText(1, 0, 'Eastern');
    table.setCellText(2, 2, 'Now filled');
    table.setCellTextDirection(1, 1, 'vert270');
    table.setCellTextFit(1, 2, 'shrink');
    table.setCellVerticalAlignment(0, 0, 'bottom');
    table.setCellVerticalAlignment(0, 1, undefined);
    table.setCellVerticalAlignment(1, 0, 'top');
    table.setCellVerticalAlignment(2, 0, 'bottom');
    table.setCellMargins(0, 0, { top: 5 });
    table.setCellMargins(2, 1, [2, 4, 6, 8]);
    table.setCellFill(0, 0, {
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 25,
    });
    table.setCellBorders(0, 0, {
      kind: 'line',
      color: { kind: 'srgb', value: 'FFFFFF' },
      width: 1,
      style: 'solid',
    });
    table.setColumnWidths([inches(1.5), inches(4.5), inches(2)]);
    table.setRowHeights([inches(0.75), inches(1.25), inches(0.5)]);
    expect(table.rowHeights).toEqual([
      inches(0.75),
      inches(1.25),
      inches(0.5),
    ]);
    expect(table.transform.height).toBe(inches(2.5));
    const explicitTransformHeight = table.transform.height;
    table.setRowHeights([0, inches(1), 0]);
    table.setTransform({ x: inches(1.5) });

    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[1]!.cells[0]!.text).toBe('Eastern');
    expect(table.rows[2]!.cells[2]!.text).toBe('Now filled');
    expect(table.rows[1]!.cells[1]!.textDirection).toBe('vert270');
    expect(table.rows[1]!.cells[2]!.textFit).toBe('shrink');
    expect(table.rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment))).toEqual([
      ['bottom', undefined, 'bottom'],
      ['top', 'middle', 'middle'],
      ['bottom', 'middle', 'middle'],
    ]);
    expect(table.rows[2]!.cells[0]!.verticalAlignment).toBe('bottom');
    expect(table.rows[2]!.cells[1]!.margins).toEqual({
      top: 2,
      right: 4,
      bottom: 6,
      left: 8,
    });
    expect(table.rows[0]!.cells[0]!.margins).toEqual({ top: 5 });
    expect(duplicateTable.rows[0]!.cells[0]!.margins).toEqual({ bottom: 9 });
    expect(table.rows[0]!.cells[0]!.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 25,
    });
    expect(table.rows[0]!.cells[0]!.borders).toEqual({
      top: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FFFFFF' },
        width: 1,
        style: 'solid',
      },
      right: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FFFFFF' },
        width: 1,
        style: 'solid',
      },
      bottom: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FFFFFF' },
        width: 1,
        style: 'solid',
      },
      left: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FFFFFF' },
        width: 1,
        style: 'solid',
      },
    });
    expect(table.transform.x).toBe(inches(1.5));
    expect(table.transform.width).toBe(inches(8));
    expect(table.columnWidths).toEqual([
      inches(1.5),
      inches(4.5),
      inches(2),
    ]);
    expect(table.rowHeights).toEqual([0, inches(1), 0]);
    expect(table.transform.height).toBe(explicitTransformHeight);
    expect(duplicateTable.columnWidths).toEqual([
      inches(2),
      inches(4),
      inches(2),
    ]);
    expect(duplicateTable.rowHeights).toEqual([
      inches(0.5),
      inches(0.75),
      inches(1),
    ]);
    expect(duplicateTable.rows).toEqual(originalRows);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    const rollbackJournal = [...document.opcPackage.mutations];
    let rolledBack: TableModel | undefined;
    expect(() =>
      document.transaction(() => {
        table.setCellFill(0, 0, { kind: 'none' });
        table.setCellBorders(0, 0, { kind: 'none' });
        table.setCellMargins(0, 0, { right: 1 });
        table.setCellVerticalAlignment(0, 0, 'top');
        table.setCellVerticalAlignment(0, 2, undefined);
        table.setColumnWidths(inches(1));
        table.setRowHeights(0);
        rolledBack = slide.addTable([[{
          text: 'temporary',
          options: {
            border: borderLine({ kind: 'scheme', value: 'accent1' }, 1, 'dash'),
            fill: { kind: 'none' },
            margin: { left: -2 },
            valign: 'middle',
          },
        }]]);
        throw new Error('restore created table state');
      }),
    ).toThrow('restore created table state');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(slide.shapes[0]).toBe(table);
    expect(table.columnWidths).toEqual([
      inches(1.5),
      inches(4.5),
      inches(2),
    ]);
    expect(table.transform.width).toBe(inches(8));
    expect(table.rowHeights).toEqual([0, inches(1), 0]);
    expect(table.transform.height).toBe(explicitTransformHeight);
    expect(table.rows[0]!.cells[0]!.borders).toEqual({
      top: borderLine({ kind: 'srgb', value: 'FFFFFF' }, 1, 'solid'),
      right: borderLine({ kind: 'srgb', value: 'FFFFFF' }, 1, 'solid'),
      bottom: borderLine({ kind: 'srgb', value: 'FFFFFF' }, 1, 'solid'),
      left: borderLine({ kind: 'srgb', value: 'FFFFFF' }, 1, 'solid'),
    });
    expect(table.rows[0]!.cells[0]!.margins).toEqual({ top: 5 });
    expect(table.rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment))).toEqual([
      ['bottom', undefined, 'bottom'],
      ['top', 'middle', 'middle'],
      ['bottom', 'middle', 'middle'],
    ]);
    expect(() => rolledBack!.rows).toThrow(ModelParseError);

    const editedMarginMatrix = table.rows.map(({ cells }) =>
      cells.map(({ margins }) => margins));
    const duplicateMarginMatrix = duplicateTable.rows.map(({ cells }) =>
      cells.map(({ margins }) => margins));
    const editedAlignmentMatrix = table.rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment));
    const duplicateAlignmentMatrix = duplicateTable.rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment));
    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    for (const [uri, expectedHash] of nonTargetHashes) {
      expect(hash(reopened.opcPackage.requirePart(uri).bytes), uri).toBe(expectedHash);
    }
    expect(reopenedTable.rows.map(({ cells }) => cells.map(({ text }) => text))).toEqual([
      ['Region', 'Revenue', 'Growth'],
      ['Eastern', '$1.2M', '12%'],
      ['West', '$980K', 'Now filled'],
    ]);
    expect(reopenedTable.rows[0]!.cells[0]!.fill).toEqual(table.rows[0]!.cells[0]!.fill);
    expect(reopenedTable.rows[0]!.cells[0]!.borders).toEqual(table.rows[0]!.cells[0]!.borders);
    expect(reopenedTable.rows[1]!.cells[1]!.textDirection).toBe('vert270');
    expect(reopenedTable.rows[1]!.cells[2]!.textFit).toBe('shrink');
    expect(reopenedTable.rows[2]!.cells[0]!.verticalAlignment).toBe('bottom');
    expect(reopenedTable.rows[2]!.cells[1]!.margins).toEqual({
      top: 2,
      right: 4,
      bottom: 6,
      left: 8,
    });
    expect(reopenedTable.rows.map(({ cells }) =>
      cells.map(({ margins }) => margins))).toEqual(editedMarginMatrix);
    expect(reopenedDuplicate.rows.map(({ cells }) =>
      cells.map(({ margins }) => margins))).toEqual(duplicateMarginMatrix);
    expect(reopenedTable.rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment))).toEqual(
      editedAlignmentMatrix,
    );
    expect(reopenedDuplicate.rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment))).toEqual(
      duplicateAlignmentMatrix,
    );
    expect(reopenedTable.transform.x).toBe(inches(1.5));
    expect(reopenedTable.transform.width).toBe(inches(8));
    expect(reopenedTable.columnWidths).toEqual([
      inches(1.5),
      inches(4.5),
      inches(2),
    ]);
    expect(reopenedTable.rowHeights).toEqual([0, inches(1), 0]);
    expect(reopenedTable.transform.height).toBe(inches(2.5));
    expect(reopenedDuplicate.columnWidths).toEqual([
      inches(2),
      inches(4),
      inches(2),
    ]);
    expect(reopenedDuplicate.rowHeights).toEqual([
      inches(0.5),
      inches(0.75),
      inches(1),
    ]);
    expect(reopenedDuplicate.rows).toEqual(originalRows);
    const reopenedTableXml = new TextDecoder().decode(
      reopened.opcPackage.requirePart(reopened.slides[0]!.partUri).bytes,
    );
    expect([...reopenedTableXml.matchAll(/<a:gridCol w="(\d+)"\/>/g)]
      .map((match) => Number(match[1]))).toEqual([
      inches(1.5),
      inches(4.5),
      inches(2),
    ]);
    expect([...reopenedTableXml.matchAll(/<a:tr h="(\d+)">/g)]
      .map((match) => Number(match[1]))).toEqual([
      0,
      inches(1),
      0,
    ]);
  });

  it('preserves created table-cell horizontal alignment through the SDK lifecycle', async () => {
    const document = PptxDocument.create({ slideSize: 'wide' });
    const slide = document.addSlide();
    const opaqueUri = '/ppt/custom/cell-horizontal-alignment.bin';
    document.opcPackage.setPart(
      opaqueUri,
      new Uint8Array([0, 17, 34, 51, 68, 85, 102, 119]),
      'application/octet-stream',
    );
    const opaqueHash = hash(document.opcPackage.requirePart(opaqueUri).bytes);
    const sourceBorderColor = { kind: 'srgb' as const, value: 'C00000' };
    const sourceBorder = {
      kind: 'line' as const,
      color: sourceBorderColor,
      width: 2,
      style: 'solid' as const,
    };
    const sourceFillColor = { kind: 'scheme' as const, value: 'accent1' };
    const sourceFill = {
      kind: 'solid' as const,
      color: sourceFillColor,
      transparency: 25,
    };
    const table = slide.addTable([
      [
        {
          text: 'Left',
          options: {
            align: 'left',
            valign: 'top',
            margin: { top: 4, left: 8 },
            border: sourceBorder,
            fill: sourceFill,
            textDirection: 'vert',
          },
        },
        {
          text: 'Center',
          options: { align: 'center', textDirection: 'vert270', valign: 'middle' },
        },
        {
          text: 'Right',
          options: { align: 'right', textDirection: 'wordArtVert', valign: 'bottom' },
        },
      ],
      [
        {
          text: 'Justify this longer sentence',
          options: { align: 'justify', textDirection: 'horz' },
        },
        {
          text: 'Undefined',
          options: { align: undefined } as unknown as AddTableCellOptions,
        },
        'Omitted',
      ],
    ], {
      name: 'Cell horizontal alignment creation',
      x: inches(1),
      y: inches(1.25),
      width: inches(6),
      height: inches(2),
      columnWidths: [inches(1.5), inches(2.5), inches(2)],
      rowHeights: [inches(0.75), inches(1.25)],
    });
    const originalTokens = ['l', 'ctr', 'r', 'just', undefined, undefined];
    expect(tableCellHorizontalAlignmentTokens(document, slide.partUri))
      .toEqual(originalTokens);
    const originalDirections = [
      'vert',
      'vert270',
      'wordArtVert',
      undefined,
      undefined,
      undefined,
    ];
    expect(table.rows.flatMap(({ cells }) =>
      cells.map(({ textDirection }) => textDirection))).toEqual(originalDirections);
    expect(table.rows[0]!.cells[0]).toMatchObject({
      text: 'Left',
      margins: { top: 4, right: 7.2, bottom: 3.6, left: 8 },
      verticalAlignment: 'top',
      fill: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 25,
      },
    });
    sourceBorderColor.value = '000000';
    sourceBorder.width = 9;
    sourceFillColor.value = 'accent6';
    sourceFill.transparency = 90;
    expect(table.rows[0]!.cells[0]!.borders?.top).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: 'C00000' },
      width: 2,
      style: 'solid',
    });
    expect(table.rows[0]!.cells[0]!.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 25,
    });

    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;
    expect(tableCellHorizontalAlignmentTokens(document, duplicate.partUri))
      .toEqual(originalTokens);
    expect(duplicateTable.rows.flatMap(({ cells }) =>
      cells.map(({ textDirection }) => textDirection))).toEqual(originalDirections);

    table.setCellText(0, 0, 'Left edited');
    table.setCellText(0, 1, 'Center edited');
    table.setCellText(0, 2, 'Right edited');
    table.setCellText(1, 0, 'Justify edited sentence');
    table.setCellMargins(0, 0, { bottom: 9 });
    table.setCellBorders(0, 1, { kind: 'none' });
    table.setCellFill(0, 2, {
      kind: 'solid',
      color: { kind: 'srgb', value: '4472C4' },
      transparency: 40,
    });
    table.setCellVerticalAlignment(1, 0, 'middle');
    table.setCellTextDirection(0, 0, 'vert270');
    table.setCellTextFit(0, 1, 'shrink');
    table.setColumnWidths([inches(2), inches(2), inches(2)]);
    table.setRowHeights([inches(1), inches(1.5)]);
    expect(tableCellHorizontalAlignmentTokens(document, slide.partUri))
      .toEqual(originalTokens);
    expect(tableCellHorizontalAlignmentTokens(document, duplicate.partUri))
      .toEqual(originalTokens);
    expect(duplicateTable.rows.map(({ cells }) => cells.map(({ text }) => text)))
      .toEqual([
        ['Left', 'Center', 'Right'],
        ['Justify this longer sentence', 'Undefined', 'Omitted'],
      ]);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const duplicateBeforeRollback = document.opcPackage
      .requirePart(duplicate.partUri).bytes.slice();
    const rollbackJournal = [...document.opcPackage.mutations];
    const slideCount = document.slides.length;
    const shapeCount = slide.shapes.length;
    let rolledBack: TableModel | undefined;
    expect(() => document.transaction(() => {
      table.setCellText(0, 0, 'Rolled back');
      table.setCellMargins(0, 1, 12);
      table.setColumnWidths(inches(3));
      rolledBack = slide.addTable([[{
        text: 'Temporary aligned cell',
        options: { align: 'center', textDirection: 'vert' },
      }]], { name: 'Temporary aligned table' });
      throw new Error('restore table cell horizontal alignment creation');
    })).toThrow('restore table cell horizontal alignment creation');
    expect(document.opcPackage.requirePart(slide.partUri).bytes)
      .toEqual(beforeRollback);
    expect(document.opcPackage.requirePart(duplicate.partUri).bytes)
      .toEqual(duplicateBeforeRollback);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(document.slides).toHaveLength(slideCount);
    expect(slide.shapes).toHaveLength(shapeCount);
    expect(slide.shapes[0]).toBe(table);
    expect(duplicate.shapes[0]).toBe(duplicateTable);
    expect(tableCellHorizontalAlignmentTokens(document, slide.partUri))
      .toEqual(originalTokens);
    expect(tableCellHorizontalAlignmentTokens(document, duplicate.partUri))
      .toEqual(originalTokens);
    expect(table.rows.flatMap(({ cells }) =>
      cells.map(({ textDirection }) => textDirection))).toEqual([
      'vert270',
      'vert270',
      'wordArtVert',
      undefined,
      undefined,
      undefined,
    ]);
    expect(duplicateTable.rows.flatMap(({ cells }) =>
      cells.map(({ textDirection }) => textDirection))).toEqual(originalDirections);
    expect(() => rolledBack!.rows).toThrow(ModelParseError);
    expect(hash(document.opcPackage.requirePart(opaqueUri).bytes)).toBe(opaqueHash);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(tableCellHorizontalAlignmentTokens(reopened, reopened.slides[0]!.partUri))
      .toEqual(originalTokens);
    expect(tableCellHorizontalAlignmentTokens(reopened, reopened.slides[1]!.partUri))
      .toEqual(originalTokens);
    expect(reopenedTable.rows.map(({ cells }) => cells.map(({ text }) => text)))
      .toEqual([
        ['Left edited', 'Center edited', 'Right edited'],
        ['Justify edited sentence', 'Undefined', 'Omitted'],
      ]);
    expect(reopenedDuplicate.rows.map(({ cells }) => cells.map(({ text }) => text)))
      .toEqual([
        ['Left', 'Center', 'Right'],
        ['Justify this longer sentence', 'Undefined', 'Omitted'],
      ]);
    expect(reopenedTable.rows.flatMap(({ cells }) =>
      cells.map(({ textDirection }) => textDirection))).toEqual([
      'vert270',
      'vert270',
      'wordArtVert',
      undefined,
      undefined,
      undefined,
    ]);
    expect(reopenedDuplicate.rows.flatMap(({ cells }) =>
      cells.map(({ textDirection }) => textDirection))).toEqual(originalDirections);
    expect(reopenedTable.rows[0]!.cells[0]!.margins).toEqual({ bottom: 9 });
    expect(reopenedTable.rows[0]!.cells[0]!.borders?.top).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: 'C00000' },
      width: 2,
      style: 'solid',
    });
    expect(reopenedTable.rows[0]!.cells[0]!.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 25,
    });
    expect(reopenedTable.rows[0]!.cells[0]!.verticalAlignment).toBe('top');
    expect(reopenedTable.rows[0]!.cells[0]!.textDirection).toBe('vert270');
    expect(reopenedTable.rows[0]!.cells[1]!.borders).toEqual({
      top: { kind: 'none' },
      right: { kind: 'none' },
      bottom: { kind: 'none' },
      left: { kind: 'none' },
    });
    expect(reopenedTable.rows[0]!.cells[1]!.textFit).toBe('shrink');
    expect(reopenedTable.rows[0]!.cells[2]!.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '4472C4' },
      transparency: 40,
    });
    expect(reopenedTable.rows[1]!.cells[0]!.verticalAlignment).toBe('middle');
    expect(reopenedTable.columnWidths).toEqual(Array(3).fill(inches(2)));
    expect(reopenedTable.rowHeights).toEqual([inches(1), inches(1.5)]);
    expect(reopenedTable.transform).toMatchObject({
      x: inches(1),
      y: inches(1.25),
      width: inches(6),
      height: inches(2.5),
    });
    expect(reopenedDuplicate.rows[0]!.cells[0]).toMatchObject({
      margins: { top: 4, right: 7.2, bottom: 3.6, left: 8 },
      verticalAlignment: 'top',
      fill: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 25,
      },
    });
    expect(reopenedDuplicate.rows[0]!.cells[0]!.borders?.top).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: 'C00000' },
      width: 2,
      style: 'solid',
    });
    expect(reopenedDuplicate.columnWidths).toEqual([
      inches(1.5),
      inches(2.5),
      inches(2),
    ]);
    expect(reopenedDuplicate.rowHeights).toEqual([
      inches(0.75),
      inches(1.25),
    ]);
    expect(reopenedDuplicate.transform).toMatchObject({
      x: inches(1),
      y: inches(1.25),
      width: inches(6),
      height: inches(2),
    });
    expect(hash(reopened.opcPackage.requirePart(opaqueUri).bytes)).toBe(opaqueHash);
    expect(validatePackage(reopened.opcPackage)
      .filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('materializes table horizontal alignment through the SDK lifecycle', async () => {
    const document = PptxDocument.create({ slideSize: 'wide' });
    const slide = document.addSlide();
    const opaqueUri = '/ppt/custom/table-horizontal-alignment.bin';
    document.opcPackage.setPart(
      opaqueUri,
      new Uint8Array([119, 102, 85, 68, 51, 34, 17, 0]),
      'application/octet-stream',
    );
    const opaqueHash = hash(document.opcPackage.requirePart(opaqueUri).bytes);
    const table = slide.addTable([[
      'Inherited string',
      { text: 'Inherited object' },
      {
        text: 'Inherited undefined',
        options: { align: undefined } as unknown as AddTableCellOptions,
      },
      { text: 'Left override', options: { align: 'left' } },
      { text: 'Right override', options: { align: 'right' } },
      { text: 'Justify override sentence', options: { align: 'justify' } },
    ]], {
      align: 'center',
      name: 'Table horizontal alignment lifecycle',
      columnWidths: inches(1.5),
      rowHeights: inches(1),
      margin: { top: 4, left: 8 },
      valign: 'middle',
    });
    const originalTokens = ['ctr', 'ctr', 'ctr', 'l', 'r', 'just'];
    expect(tableCellHorizontalAlignmentTokens(document, slide.partUri))
      .toEqual(originalTokens);
    expect(table.rows[0]!.cells.map(({ verticalAlignment }) => verticalAlignment))
      .toEqual(Array(6).fill('middle'));
    expect(table.rows[0]!.cells[0]!.margins).toEqual({
      top: 4,
      right: 7.2,
      bottom: 3.6,
      left: 8,
    });

    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;
    expect(tableCellHorizontalAlignmentTokens(document, duplicate.partUri))
      .toEqual(originalTokens);

    table.setCellText(0, 0, 'Inherited edited');
    table.setCellMargins(0, 1, { bottom: 9 });
    table.setCellVerticalAlignment(0, 2, 'bottom');
    table.setCellTextDirection(0, 3, 'vert270');
    table.setCellTextFit(0, 4, 'shrink');
    table.setCellBorders(0, 0, { kind: 'none' });
    table.setCellFill(0, 5, {
      kind: 'solid',
      color: { kind: 'srgb', value: 'FFF2CC' },
    });
    table.setColumnWidths(inches(1.25));
    table.setRowHeights(inches(1.25));
    expect(tableCellHorizontalAlignmentTokens(document, slide.partUri))
      .toEqual(originalTokens);
    expect(tableCellHorizontalAlignmentTokens(document, duplicate.partUri))
      .toEqual(originalTokens);
    expect(duplicateTable.rows[0]!.cells[0]!.text).toBe('Inherited string');

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const duplicateBeforeRollback = document.opcPackage
      .requirePart(duplicate.partUri).bytes.slice();
    const rollbackJournal = [...document.opcPackage.mutations];
    const shapeCount = slide.shapes.length;
    let rolledBack: TableModel | undefined;
    expect(() => document.transaction(() => {
      table.setCellText(0, 0, 'Rolled back');
      rolledBack = slide.addTable([['Temporary']], { align: 'right' });
      throw new Error('restore table horizontal alignment');
    })).toThrow('restore table horizontal alignment');
    expect(document.opcPackage.requirePart(slide.partUri).bytes)
      .toEqual(beforeRollback);
    expect(document.opcPackage.requirePart(duplicate.partUri).bytes)
      .toEqual(duplicateBeforeRollback);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(slide.shapes).toHaveLength(shapeCount);
    expect(slide.shapes[0]).toBe(table);
    expect(duplicate.shapes[0]).toBe(duplicateTable);
    expect(() => rolledBack!.rows).toThrow(ModelParseError);
    expect(hash(document.opcPackage.requirePart(opaqueUri).bytes)).toBe(opaqueHash);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(tableCellHorizontalAlignmentTokens(reopened, reopened.slides[0]!.partUri))
      .toEqual(originalTokens);
    expect(tableCellHorizontalAlignmentTokens(reopened, reopened.slides[1]!.partUri))
      .toEqual(originalTokens);
    expect(reopenedTable.rows[0]!.cells.map(({ text }) => text)).toEqual([
      'Inherited edited',
      'Inherited object',
      'Inherited undefined',
      'Left override',
      'Right override',
      'Justify override sentence',
    ]);
    expect(reopenedTable.rows[0]!.cells[1]!.margins).toEqual({ bottom: 9 });
    expect(reopenedTable.rows[0]!.cells[2]!.verticalAlignment).toBe('bottom');
    expect(reopenedTable.rows[0]!.cells[3]!.textDirection).toBe('vert270');
    expect(reopenedTable.rows[0]!.cells[4]!.textFit).toBe('shrink');
    expect(reopenedTable.rows[0]!.cells[0]!.borders).toEqual({
      top: { kind: 'none' },
      right: { kind: 'none' },
      bottom: { kind: 'none' },
      left: { kind: 'none' },
    });
    expect(reopenedTable.rows[0]!.cells[5]!.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'FFF2CC' },
    });
    expect(reopenedTable.columnWidths).toEqual(Array(6).fill(inches(1.25)));
    expect(reopenedTable.rowHeights).toEqual([inches(1.25)]);
    expect(reopenedDuplicate.rows[0]!.cells[0]!.text).toBe('Inherited string');
    expect(reopenedDuplicate.columnWidths).toEqual(Array(6).fill(inches(1.5)));
    expect(reopenedDuplicate.rowHeights).toEqual([inches(1)]);
    expect(hash(reopened.opcPackage.requirePart(opaqueUri).bytes)).toBe(opaqueHash);
    expect(validatePackage(reopened.opcPackage)
      .filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('materializes table text direction through the SDK lifecycle', async () => {
    const document = PptxDocument.create({ slideSize: 'wide' });
    const slide = document.addSlide();
    const opaqueUri = '/ppt/custom/table-text-direction.bin';
    document.opcPackage.setPart(
      opaqueUri,
      new Uint8Array([13, 26, 39, 52, 65, 78, 91, 104]),
      'application/octet-stream',
    );
    const opaqueHash = hash(document.opcPackage.requirePart(opaqueUri).bytes);
    const table = slide.addTable([
      [
        'Inherited string',
        {
          text: 'Inherited undefined',
          options: { textDirection: undefined } as unknown as AddTableCellOptions,
        },
        { text: 'Horizontal override', options: { textDirection: 'horz' } },
      ],
      [
        { text: 'Vertical override', options: { textDirection: 'vert' } },
        { text: 'Stacked override', options: { textDirection: 'wordArtVert' } },
        { text: 'Inherited object', options: {} },
      ],
    ], {
      name: 'Table text direction lifecycle',
      textDirection: 'vert270',
      columnWidths: inches(2),
      rowHeights: inches(1),
      margin: { top: 4, left: 8 },
      valign: 'middle',
    });
    const original = [
      'vert270',
      'vert270',
      undefined,
      'vert',
      'wordArtVert',
      'vert270',
    ];
    const directDirections = (target: PptxDocument, partUri: string) => {
      const xml = new TextDecoder().decode(target.opcPackage.requirePart(partUri).bytes);
      return [...xml.matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g)]
        .map((match) => match[0]!
          .match(/<a:tcPr[^>]*\svert="([^"]+)"/)?.[1]);
    };
    expect(table.rows.flatMap(({ cells }) =>
      cells.map(({ textDirection }) => textDirection))).toEqual(original);
    expect(directDirections(document, slide.partUri)).toEqual(original);

    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;
    expect(duplicateTable.rows.flatMap(({ cells }) =>
      cells.map(({ textDirection }) => textDirection))).toEqual(original);
    expect(directDirections(document, duplicate.partUri)).toEqual(original);

    table.setCellTextDirection(0, 0, undefined);
    table.setCellTextDirection(0, 1, 'wordArtVert');
    table.setCellText(1, 2, 'Inherited object edited');
    table.setCellMargins(0, 0, { bottom: 9 });
    table.setCellVerticalAlignment(0, 1, 'bottom');
    table.setCellBorders(1, 0, { kind: 'none' });
    table.setCellFill(1, 1, {
      kind: 'solid',
      color: { kind: 'srgb', value: 'FFF2CC' },
    });
    table.setCellTextFit(1, 2, 'shrink');
    table.setColumnWidths(inches(1.75));
    table.setRowHeights(inches(1.25));
    const edited = [
      undefined,
      'wordArtVert',
      undefined,
      'vert',
      'wordArtVert',
      'vert270',
    ];
    expect(table.rows.flatMap(({ cells }) =>
      cells.map(({ textDirection }) => textDirection))).toEqual(edited);
    expect(directDirections(document, slide.partUri)).toEqual(edited);
    expect(duplicateTable.rows.flatMap(({ cells }) =>
      cells.map(({ textDirection }) => textDirection))).toEqual(original);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const duplicateBeforeRollback = document.opcPackage
      .requirePart(duplicate.partUri).bytes.slice();
    const rollbackJournal = [...document.opcPackage.mutations];
    const shapeCount = slide.shapes.length;
    let rolledBack: TableModel | undefined;
    expect(() => document.transaction(() => {
      table.setCellTextDirection(0, 0, 'vert');
      table.setCellText(0, 1, 'Rolled back');
      rolledBack = slide.addTable([['Temporary']], {
        textDirection: 'wordArtVert',
      });
      throw new Error('restore table text direction');
    })).toThrow('restore table text direction');
    expect(document.opcPackage.requirePart(slide.partUri).bytes)
      .toEqual(beforeRollback);
    expect(document.opcPackage.requirePart(duplicate.partUri).bytes)
      .toEqual(duplicateBeforeRollback);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(slide.shapes).toHaveLength(shapeCount);
    expect(slide.shapes[0]).toBe(table);
    expect(duplicate.shapes[0]).toBe(duplicateTable);
    expect(() => rolledBack!.rows).toThrow(ModelParseError);
    expect(table.rows.flatMap(({ cells }) =>
      cells.map(({ textDirection }) => textDirection))).toEqual(edited);
    expect(hash(document.opcPackage.requirePart(opaqueUri).bytes)).toBe(opaqueHash);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(reopenedTable.rows.flatMap(({ cells }) =>
      cells.map(({ textDirection }) => textDirection))).toEqual(edited);
    expect(directDirections(reopened, reopened.slides[0]!.partUri)).toEqual(edited);
    expect(reopenedDuplicate.rows.flatMap(({ cells }) =>
      cells.map(({ textDirection }) => textDirection))).toEqual(original);
    expect(directDirections(reopened, reopened.slides[1]!.partUri)).toEqual(original);
    expect(reopenedTable.rows[1]!.cells[2]!.text).toBe('Inherited object edited');
    expect(reopenedTable.rows[0]!.cells[0]!.margins).toEqual({ bottom: 9 });
    expect(reopenedTable.rows[0]!.cells[1]!.verticalAlignment).toBe('bottom');
    expect(reopenedTable.rows[1]!.cells[0]!.borders).toEqual({
      top: { kind: 'none' },
      right: { kind: 'none' },
      bottom: { kind: 'none' },
      left: { kind: 'none' },
    });
    expect(reopenedTable.rows[1]!.cells[1]!.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'FFF2CC' },
    });
    expect(reopenedTable.rows[1]!.cells[2]!.textFit).toBe('shrink');
    expect(reopenedTable.columnWidths).toEqual(Array(3).fill(inches(1.75)));
    expect(reopenedTable.rowHeights).toEqual(Array(2).fill(inches(1.25)));
    expect(reopenedDuplicate.columnWidths).toEqual(Array(3).fill(inches(2)));
    expect(reopenedDuplicate.rowHeights).toEqual(Array(2).fill(inches(1)));
    expect(hash(reopened.opcPackage.requirePart(opaqueUri).bytes)).toBe(opaqueHash);
    expect(validatePackage(reopened.opcPackage)
      .filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('materializes public table margins through duplicate, rollback, and reopen', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const table = slide.addTable([[
      'Inherited string',
      { text: 'Partial override', options: { margin: { bottom: 12 } } },
      { text: 'Zero override', options: { margin: 0 } },
      { text: 'Tuple override', options: { margin: [1, 2, 3, 4] } },
    ]], {
      name: 'SDK table margin lifecycle',
      margin: { top: 9, left: 18 },
      columnWidths: inches(2),
      rowHeights: inches(1),
    });
    const original = table.rows[0]!.cells.map(({ margins }) => margins);
    expect(original).toEqual([
      { top: 9, right: 7.2, bottom: 3.6, left: 18 },
      { top: 9, right: 7.2, bottom: 12, left: 18 },
      { top: 0, right: 0, bottom: 0, left: 0 },
      { top: 1, right: 2, bottom: 3, left: 4 },
    ]);

    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;
    expect(duplicateTable.rows[0]!.cells.map(({ margins }) => margins))
      .toEqual(original);

    table.setCellMargins(0, 0, undefined);
    table.setCellMargins(0, 1, { right: 5 });
    const edited = table.rows[0]!.cells.map(({ margins }) => margins);
    expect(edited).toEqual([
      undefined,
      { right: 5 },
      { top: 0, right: 0, bottom: 0, left: 0 },
      { top: 1, right: 2, bottom: 3, left: 4 },
    ]);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const rollbackJournal = [...document.opcPackage.mutations];
    expect(() => document.transaction(() => {
      table.setCellMargins(0, 2, { left: 11 });
      slide.addTable([['Temporary']], { margin: 6 });
      throw new Error('restore table margin defaults');
    })).toThrow('restore table margin defaults');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(slide.shapes).toHaveLength(1);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[0]!.cells.map(({ margins }) => margins)).toEqual(edited);
    expect(duplicateTable.rows[0]!.cells.map(({ margins }) => margins))
      .toEqual(original);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(reopenedTable.rows[0]!.cells.map(({ margins }) => margins))
      .toEqual(edited);
    expect(reopenedDuplicate.rows[0]!.cells.map(({ margins }) => margins))
      .toEqual(original);
    expect(reopenedTable.columnWidths).toEqual(Array(4).fill(inches(2)));
    expect(reopenedTable.rowHeights).toEqual([inches(1)]);
    expect(reopenedDuplicate.columnWidths).toEqual(Array(4).fill(inches(2)));
    expect(reopenedDuplicate.rowHeights).toEqual([inches(1)]);
  });

  it('materializes public table fills through duplicate, rollback, and reopen', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const sourceColor = {
      kind: 'scheme' as const,
      value: 'accent1' as 'accent1' | 'accent6',
    };
    const sourceFill = {
      kind: 'solid' as const,
      color: sourceColor,
      transparency: 33.3334,
    };
    const table = slide.addTable([[
      'Inherited string',
      { text: 'Inherited object', options: {} },
      { text: 'None override', options: { fill: { kind: 'none' } } },
      { text: 'Solid override', options: { fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFFF00' },
        transparency: 25,
      } } },
    ]], {
      name: 'SDK table fill lifecycle',
      fill: sourceFill,
      columnWidths: inches(2),
      rowHeights: inches(1),
    });
    const original = table.rows[0]!.cells.map(({ fill }) => fill);
    expect(original).toEqual([
      {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 33.333,
      },
      {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 33.333,
      },
      { kind: 'none' },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFFF00' },
        transparency: 25,
      },
    ]);
    sourceColor.value = 'accent6';
    sourceFill.transparency = 1;
    expect(table.rows[0]!.cells.map(({ fill }) => fill)).toEqual(original);

    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;
    expect(duplicateTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(original);

    table.setCellFill(0, 0, undefined);
    table.setCellFill(0, 1, {
      kind: 'solid',
      color: { kind: 'srgb', value: '00FF00' },
      transparency: 0,
    });
    const edited = table.rows[0]!.cells.map(({ fill }) => fill);
    expect(edited).toEqual([
      undefined,
      {
        kind: 'solid',
        color: { kind: 'srgb', value: '00FF00' },
        transparency: 0,
      },
      { kind: 'none' },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFFF00' },
        transparency: 25,
      },
    ]);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const rollbackJournal = [...document.opcPackage.mutations];
    let rolledBack: TableModel | undefined;
    expect(() => document.transaction(() => {
      table.setCellFill(0, 2, {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
      });
      rolledBack = slide.addTable([['Temporary']], { fill: { kind: 'none' } });
      throw new Error('restore table fill defaults');
    })).toThrow('restore table fill defaults');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(slide.shapes).toHaveLength(1);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[0]!.cells.map(({ fill }) => fill)).toEqual(edited);
    expect(duplicateTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(original);
    expect(() => rolledBack!.rows).toThrow(ModelParseError);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(reopenedTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(edited);
    expect(reopenedDuplicate.rows[0]!.cells.map(({ fill }) => fill)).toEqual(original);
    expect(reopenedTable.columnWidths).toEqual(Array(4).fill(inches(2)));
    expect(reopenedTable.rowHeights).toEqual([inches(1)]);
    expect(reopenedDuplicate.columnWidths).toEqual(Array(4).fill(inches(2)));
    expect(reopenedDuplicate.rowHeights).toEqual([inches(1)]);
    expect(validatePackage(reopened.opcPackage)
      .filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('materializes public table borders through duplicate, rollback, and reopen', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const sourceColor = {
      kind: 'scheme' as const,
      value: 'accent1' as 'accent1' | 'accent6',
    };
    const sourceBorder = {
      kind: 'line' as const,
      color: sourceColor,
      width: 1.500004,
      style: 'dash' as const,
    };
    const table = slide.addTable([[
      'Inherited string',
      { text: 'Inherited empty', options: { border: {} } },
      { text: 'Partial override', options: { border: {
        bottom: {
          kind: 'line',
          color: { kind: 'srgb', value: '70AD47' },
          width: 3,
          style: 'solid',
        },
      } } },
      { text: 'None override', options: { border: { kind: 'none' } } },
    ]], {
      name: 'SDK table border lifecycle',
      border: sourceBorder,
      columnWidths: inches(2),
      rowHeights: inches(1),
    });
    const tableDefault = {
      top: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent1' },
        width: 1.5,
        style: 'dash',
      },
      right: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent1' },
        width: 1.5,
        style: 'dash',
      },
      bottom: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent1' },
        width: 1.5,
        style: 'dash',
      },
      left: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent1' },
        width: 1.5,
        style: 'dash',
      },
    };
    const noBorders = {
      top: { kind: 'none' },
      right: { kind: 'none' },
      bottom: { kind: 'none' },
      left: { kind: 'none' },
    };
    const original = [
      tableDefault,
      tableDefault,
      {
        top: { kind: 'none' },
        right: { kind: 'none' },
        bottom: {
          kind: 'line',
          color: { kind: 'srgb', value: '70AD47' },
          width: 3,
          style: 'solid',
        },
        left: { kind: 'none' },
      },
      noBorders,
    ];
    expect(table.rows[0]!.cells.map(({ borders }) => borders)).toEqual(original);
    sourceColor.value = 'accent6';
    sourceBorder.width = 9;
    expect(table.rows[0]!.cells.map(({ borders }) => borders)).toEqual(original);

    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;
    expect(duplicateTable.rows[0]!.cells.map(({ borders }) => borders))
      .toEqual(original);
    const nonTargetHashes = new Map(
      document.opcPackage.parts
        .filter(({ uri }) => uri !== slide.partUri)
        .map(({ uri, bytes }) => [uri, hash(bytes)]),
    );

    table.setCellBorders(0, 0, undefined);
    table.setCellBorders(0, 1, {
      right: {
        kind: 'line',
        color: { kind: 'srgb', value: '00FF00' },
        width: 0,
        style: 'solid',
      },
    });
    const edited = [
      undefined,
      {
        right: {
          kind: 'line',
          color: { kind: 'srgb', value: '00FF00' },
          width: 0,
          style: 'solid',
        },
      },
      original[2],
      noBorders,
    ];
    expect(table.rows[0]!.cells.map(({ borders }) => borders)).toEqual(edited);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const rollbackJournal = [...document.opcPackage.mutations];
    let rolledBack: TableModel | undefined;
    expect(() => document.transaction(() => {
      table.setCellBorders(0, 2, {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        width: 2,
      });
      rolledBack = slide.addTable([['Temporary']], { border: { kind: 'none' } });
      throw new Error('restore table border defaults');
    })).toThrow('restore table border defaults');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(slide.shapes).toHaveLength(1);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[0]!.cells.map(({ borders }) => borders)).toEqual(edited);
    expect(duplicateTable.rows[0]!.cells.map(({ borders }) => borders))
      .toEqual(original);
    expect(() => rolledBack!.rows).toThrow(ModelParseError);
    for (const [uri, expectedHash] of nonTargetHashes) {
      expect(hash(document.opcPackage.requirePart(uri).bytes), uri).toBe(expectedHash);
    }

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(reopenedTable.rows[0]!.cells.map(({ borders }) => borders)).toEqual(edited);
    expect(reopenedDuplicate.rows[0]!.cells.map(({ borders }) => borders))
      .toEqual(original);
    expect(reopenedTable.columnWidths).toEqual(Array(4).fill(inches(2)));
    expect(reopenedTable.rowHeights).toEqual([inches(1)]);
    expect(reopenedDuplicate.columnWidths).toEqual(Array(4).fill(inches(2)));
    expect(reopenedDuplicate.rowHeights).toEqual([inches(1)]);
    for (const [uri, expectedHash] of nonTargetHashes) {
      expect(hash(reopened.opcPackage.requirePart(uri).bytes), uri).toBe(expectedHash);
    }
    expect(validatePackage(reopened.opcPackage)
      .filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('rejects invalid public table border creation before mutation', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const table = slide.addTable([['Existing']]);
    const existingBorders = table.rows[0]!.cells[0]!.borders;
    const beforeParts = new Map(
      document.opcPackage.parts.map(({ uri, bytes }) => [uri, bytes.slice()]),
    );
    const journal = [...document.opcPackage.mutations];
    let sdkBorderGetterCalls = 0;
    const accessorOptions = {};
    Object.defineProperty(accessorOptions, 'border', {
      get() {
        sdkBorderGetterCalls += 1;
        return { kind: 'none' };
      },
      enumerable: true,
    });
    const accessorBorder = {};
    Object.defineProperty(accessorBorder, 'kind', {
      get() {
        sdkBorderGetterCalls += 1;
        return 'none';
      },
      enumerable: true,
    });
    const accessorNamed = {};
    Object.defineProperty(accessorNamed, 'top', {
      get() {
        sdkBorderGetterCalls += 1;
        return { kind: 'none' };
      },
      enumerable: true,
    });
    const accessorTuple = [undefined, undefined, undefined, undefined];
    Object.defineProperty(accessorTuple, '0', {
      get() {
        sdkBorderGetterCalls += 1;
        return { kind: 'none' };
      },
      enumerable: true,
    });
    const accessorColor = { kind: 'srgb' };
    Object.defineProperty(accessorColor, 'value', {
      get() {
        sdkBorderGetterCalls += 1;
        return 'FF0000';
      },
      enumerable: true,
    });
    class SdkBorderClass { kind = 'none'; }
    class SdkColorClass { kind = 'srgb'; value = 'FF0000'; }
    const sparse = Array(4);
    sparse[0] = { kind: 'none' };
    const extraTuple = Object.assign(
      [{ kind: 'none' }, undefined, undefined, undefined],
      { extra: true },
    );
    const invalidBorders: unknown[] = [
      null,
      false,
      true,
      'FF0000',
      [],
      [{ kind: 'none' }],
      [{ kind: 'none' }, undefined, undefined],
      [{ kind: 'none' }, undefined, undefined, undefined, undefined],
      sparse,
      extraTuple,
      accessorBorder,
      accessorNamed,
      accessorTuple,
      new SdkBorderClass(),
      Object.create({ kind: 'none' }),
      { top: undefined, extra: true },
      { kind: 'none', width: 1 },
      { kind: 'none', [Symbol('border')]: true },
      { kind: 'unknown' },
      { kind: 'line' },
      { kind: 'line', color: null, width: 1 },
      { kind: 'line', color: accessorColor, width: 1 },
      { kind: 'line', color: new SdkColorClass(), width: 1 },
      { kind: 'line', color: Object.create({
        kind: 'srgb', value: 'FF0000',
      }), width: 1 },
      { kind: 'line', color: { kind: 'srgb', value: 'FFF' }, width: 1 },
      { kind: 'line', color: { kind: 'srgb', value: 'GG0000' }, width: 1 },
      { kind: 'line', color: { kind: 'scheme', value: 'unknown' }, width: 1 },
      { kind: 'line', color: {
        kind: 'srgb', value: 'FF0000', extra: true,
      }, width: 1 },
      { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: -0.001 },
      { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 1584.001 },
      { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: Number.NaN },
      { kind: 'line', color: {
        kind: 'srgb', value: 'FF0000',
      }, width: Number.POSITIVE_INFINITY },
      { kind: 'line', color: {
        kind: 'srgb', value: 'FF0000',
      }, width: 1, style: 'dot' },
      Symbol('table border'),
    ];

    expect(() => slide.addTable(
      [['Accessor table']],
      accessorOptions as AddTableOptions,
    )).toThrow();
    for (const border of invalidBorders) {
      expect(() => slide.addTable([[{
        text: 'Invalid cell',
        options: { border } as unknown as AddTableCellOptions,
      }]])).toThrow();
      expect(() => slide.addTable(
        [['Invalid table']],
        { border } as unknown as AddTableOptions,
      )).toThrow();
    }

    expect(sdkBorderGetterCalls).toBe(0);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(document.slides).toHaveLength(1);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes).toHaveLength(1);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[0]!.cells[0]!.text).toBe('Existing');
    expect(table.rows[0]!.cells[0]!.borders).toEqual(existingBorders);
    expect(document.opcPackage.parts.map(({ uri }) => uri))
      .toEqual([...beforeParts.keys()]);
    for (const [uri, bytes] of beforeParts) {
      expect(document.opcPackage.requirePart(uri).bytes, uri).toEqual(bytes);
    }
  });

  it('rejects invalid public table fill creation before mutation', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const table = slide.addTable([['Existing']]);
    const before = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const journal = [...document.opcPackage.mutations];
    let sdkFillGetterCalls = 0;
    const accessorOptions = {};
    Object.defineProperty(accessorOptions, 'fill', {
      get() {
        sdkFillGetterCalls += 1;
        return { kind: 'none' };
      },
      enumerable: true,
    });
    const accessorFill = {};
    Object.defineProperty(accessorFill, 'kind', {
      get() {
        sdkFillGetterCalls += 1;
        return 'none';
      },
      enumerable: true,
    });
    const accessorColor = { kind: 'srgb' };
    Object.defineProperty(accessorColor, 'value', {
      get() {
        sdkFillGetterCalls += 1;
        return 'FF0000';
      },
      enumerable: true,
    });
    class SdkFillClass {
      kind = 'none';
    }
    class SdkColorClass {
      kind = 'srgb';
      value = 'FF0000';
    }
    const invalidFills: unknown[] = [
      null,
      false,
      true,
      'FF0000',
      [],
      {},
      accessorFill,
      new SdkFillClass(),
      Object.create({ kind: 'none' }),
      { kind: 'none', color: { kind: 'srgb', value: 'FF0000' } },
      { kind: 'none', transparency: 0 },
      { kind: 'none', extra: true },
      { kind: 'none', [Symbol('fill')]: true },
      { kind: 'unknown' },
      { kind: 'solid' },
      { kind: 'solid', color: null },
      { kind: 'solid', color: accessorColor },
      { kind: 'solid', color: new SdkColorClass() },
      { kind: 'solid', color: Object.create({ kind: 'srgb', value: 'FF0000' }) },
      { kind: 'solid', color: { kind: 'srgb', value: 'FFF' } },
      { kind: 'solid', color: { kind: 'srgb', value: 'GG0000' } },
      { kind: 'solid', color: { kind: 'scheme', value: 'unknown' } },
      { kind: 'solid', color: { kind: 'srgb', value: 'FF0000', extra: true } },
      { kind: 'solid', color: {
        kind: 'srgb',
        value: 'FF0000',
        [Symbol('color')]: true,
      } },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
        transparency: -0.001,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
        transparency: 100.001,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
        transparency: Number.NaN,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
        transparency: Number.POSITIVE_INFINITY,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
        extra: true,
      },
      Symbol('table fill'),
    ];

    expect(() => slide.addTable(
      [['Accessor table']],
      accessorOptions as AddTableOptions,
    )).toThrow();
    for (const fill of invalidFills) {
      expect(() => slide.addTable([[{
        text: 'Invalid cell',
        options: { fill } as unknown as AddTableCellOptions,
      }]])).toThrow();
      expect(() => slide.addTable(
        [['Invalid table']],
        { fill } as unknown as AddTableOptions,
      )).toThrow();
    }

    expect(sdkFillGetterCalls).toBe(0);
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(document.slides).toHaveLength(1);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes).toHaveLength(1);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[0]!.cells[0]!.text).toBe('Existing');
    expect(table.rows[0]!.cells[0]!.fill).toBeUndefined();
  });

  it('rejects invalid public table margin creation before mutation', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const table = slide.addTable([['Existing']]);
    const before = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const journal = [...document.opcPackage.mutations];
    let sdkMarginGetterCalls = 0;
    const accessorOptions = {};
    Object.defineProperty(accessorOptions, 'margin', {
      get() {
        sdkMarginGetterCalls += 1;
        return 1;
      },
      enumerable: true,
    });
    const accessorNamed = { left: 1 };
    Object.defineProperty(accessorNamed, 'top', {
      get() {
        sdkMarginGetterCalls += 1;
        return 2;
      },
      enumerable: true,
    });
    const accessorTuple = [1, 2, 3, 4];
    Object.defineProperty(accessorTuple, '1', {
      get() {
        sdkMarginGetterCalls += 1;
        return 2;
      },
      enumerable: true,
    });
    class SdkMarginClass {
      top = 1;
    }
    const sparseTuple = [1, 2, 3, 4];
    delete sparseTuple[2];
    const invalidMargins: unknown[] = [
      null,
      false,
      '1',
      Number.NaN,
      Number.POSITIVE_INFINITY,
      Number.NEGATIVE_INFINITY,
      accessorNamed,
      accessorTuple,
      new SdkMarginClass(),
      Object.create({ top: 1 }),
      { top: 1, [Symbol('margin')]: 2 },
      Object.assign([1, 2, 3, 4], { extra: true }),
      sparseTuple,
      [1, 2, 3],
      [1, 2, 3, 4, 5],
      2_147_483_648 / 12_700 + 1,
    ];

    expect(() => slide.addTable(
      [['Accessor table']],
      accessorOptions as AddTableOptions,
    )).toThrow();
    for (const margin of invalidMargins) {
      expect(() => slide.addTable([[{
        text: 'Invalid',
        options: { margin } as unknown as AddTableCellOptions,
      }]])).toThrow();
      expect(() => slide.addTable(
        [['Invalid table']],
        { margin } as unknown as AddTableOptions,
      )).toThrow();
    }

    expect(sdkMarginGetterCalls).toBe(0);
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(document.slides).toHaveLength(1);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes).toHaveLength(1);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[0]!.cells[0]!.text).toBe('Existing');
    expect(table.rows[0]!.cells[0]!.margins).toEqual({
      top: 3.6,
      right: 7.2,
      bottom: 3.6,
      left: 7.2,
    });
  });

  it('rejects invalid public table valign creation before mutation', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const table = slide.addTable([['Existing']]);
    const before = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const journal = [...document.opcPackage.mutations];
    let sdkValignGetterCalls = 0;
    const accessorOptions = {};
    Object.defineProperty(accessorOptions, 'valign', {
      get() {
        sdkValignGetterCalls += 1;
        return 'top';
      },
      enumerable: true,
    });
    const invalidValigns: unknown[] = [
      null,
      false,
      true,
      0,
      '',
      'Top',
      ' top ',
      'mid',
      'center',
      'just',
      'dist',
      'distributed',
      [],
      {},
      Symbol('top'),
    ];

    expect(() => slide.addTable([[{
      text: 'Accessor',
      options: accessorOptions as AddTableCellOptions,
    }]])).toThrow();
    expect(() => slide.addTable(
      [['Accessor table']],
      accessorOptions as AddTableOptions,
    )).toThrow();
    for (const valign of invalidValigns) {
      expect(() => slide.addTable([[{
        text: 'Invalid',
        options: { valign } as unknown as AddTableCellOptions,
      }]])).toThrow(TypeError);
      expect(() => slide.addTable(
        [['Invalid table']],
        { valign } as unknown as AddTableOptions,
      )).toThrow(TypeError);
    }

    expect(sdkValignGetterCalls).toBe(0);
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(document.slides).toHaveLength(1);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes).toHaveLength(1);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[0]!.cells[0]!.text).toBe('Existing');
    expect(table.rows[0]!.cells[0]!.verticalAlignment).toBeUndefined();
  });

  it('rejects invalid public table-cell horizontal alignment before mutation', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const table = slide.addTable([[{
      text: 'Existing',
      options: { align: 'center' },
    }]]);
    const beforeParts = new Map(
      document.opcPackage.parts.map(({ uri, bytes }) => [uri, bytes.slice()]),
    );
    const journal = [...document.opcPackage.mutations];
    const slideCount = document.slides.length;
    const shapeCount = slide.shapes.length;
    const existingTokens = tableCellHorizontalAlignmentTokens(document, slide.partUri);
    let sdkAlignGetterCalls = 0;
    const accessorOptions = {};
    Object.defineProperty(accessorOptions, 'align', {
      get() {
        sdkAlignGetterCalls += 1;
        return 'center';
      },
      enumerable: true,
    });
    const invalidAlignments: unknown[] = [
      null,
      false,
      true,
      0,
      '',
      'Left',
      ' center ',
      'l',
      'ctr',
      'r',
      'just',
      'dist',
      'thaiDist',
      'justLow',
      [],
      {},
      Symbol('center'),
    ];

    expect(() => slide.addTable([[{
      text: 'Accessor',
      options: accessorOptions as AddTableCellOptions,
    }]])).toThrow(TypeError);
    for (const align of invalidAlignments) {
      expect(() => slide.addTable([[{
        text: 'Invalid',
        options: { align } as unknown as AddTableCellOptions,
      }]])).toThrow(TypeError);
    }

    expect(sdkAlignGetterCalls).toBe(0);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(document.slides).toHaveLength(slideCount);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes).toHaveLength(shapeCount);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[0]!.cells[0]!.text).toBe('Existing');
    expect(tableCellHorizontalAlignmentTokens(document, slide.partUri))
      .toEqual(existingTokens);
    expect(document.opcPackage.parts.map(({ uri }) => uri))
      .toEqual([...beforeParts.keys()]);
    for (const [uri, bytes] of beforeParts) {
      expect(document.opcPackage.requirePart(uri).bytes, uri).toEqual(bytes);
    }
  });

  it('rejects invalid public table and table-cell text direction creation before mutation', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const table = slide.addTable([[{
      text: 'Existing',
      options: { textDirection: 'vert' },
    }]]);
    const beforeParts = new Map(
      document.opcPackage.parts.map(({ uri, bytes }) => [uri, bytes.slice()]),
    );
    const journal = [...document.opcPackage.mutations];
    let getterCalls = 0;
    const accessorOptions = {};
    Object.defineProperty(accessorOptions, 'textDirection', {
      get() {
        getterCalls += 1;
        return 'vert';
      },
      enumerable: true,
      configurable: true,
    });
    const invalidDirections: unknown[] = [
      null,
      false,
      true,
      0,
      '',
      'Vert',
      ' vert ',
      'eaVert',
      'mongolianVert',
      'wordArtVertRtl',
      [],
      {},
      Symbol('table cell text direction'),
    ];

    expect(() => slide.addTable([[{
      text: 'Accessor',
      options: accessorOptions as AddTableCellOptions,
    }]])).toThrow(TypeError);
    expect(() => slide.addTable(
      [['Accessor table']],
      accessorOptions as AddTableOptions,
    )).toThrow(TypeError);
    for (const textDirection of invalidDirections) {
      expect(() => slide.addTable([[{
        text: 'Invalid',
        options: { textDirection } as unknown as AddTableCellOptions,
      }]])).toThrow(TypeError);
      expect(() => slide.addTable(
        [['Invalid table']],
        { textDirection } as unknown as AddTableOptions,
      )).toThrow(TypeError);
    }

    expect(getterCalls).toBe(0);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(slide.shapes).toHaveLength(1);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[0]!.cells[0]!.textDirection).toBe('vert');
    expect(document.opcPackage.parts.map(({ uri }) => uri))
      .toEqual([...beforeParts.keys()]);
    for (const [uri, bytes] of beforeParts) {
      expect(document.opcPackage.requirePart(uri).bytes, uri).toEqual(bytes);
    }
  });

  it('rejects invalid public table horizontal alignment before mutation', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const table = slide.addTable([
      ['Inherited', { text: 'Right', options: { align: 'right' } }],
    ], { align: 'center' });
    const beforeParts = new Map(
      document.opcPackage.parts.map(({ uri, bytes }) => [uri, bytes.slice()]),
    );
    const journal = [...document.opcPackage.mutations];
    const slideCount = document.slides.length;
    const shapeCount = slide.shapes.length;
    const existingTokens = tableCellHorizontalAlignmentTokens(document, slide.partUri);
    let sdkTableAlignGetterCalls = 0;
    const accessorOptions = {};
    Object.defineProperty(accessorOptions, 'align', {
      get() {
        sdkTableAlignGetterCalls += 1;
        return 'center';
      },
      enumerable: true,
      configurable: true,
    });
    const invalidAlignments: unknown[] = [
      null,
      false,
      true,
      0,
      '',
      'Left',
      ' center ',
      'l',
      'ctr',
      'r',
      'just',
      'dist',
      'thaiDist',
      'justLow',
      [],
      {},
      Symbol('center'),
    ];

    expect(() => slide.addTable(
      [['Accessor']],
      accessorOptions as AddTableOptions,
    )).toThrow(TypeError);
    for (const align of invalidAlignments) {
      expect(() => slide.addTable(
        [['Invalid']],
        { align } as unknown as AddTableOptions,
      )).toThrow(TypeError);
    }

    expect(sdkTableAlignGetterCalls).toBe(0);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(document.slides).toHaveLength(slideCount);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes).toHaveLength(shapeCount);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[0]!.cells.map(({ text }) => text)).toEqual([
      'Inherited',
      'Right',
    ]);
    expect(tableCellHorizontalAlignmentTokens(document, slide.partUri))
      .toEqual(existingTokens);
    expect(document.opcPackage.parts.map(({ uri }) => uri))
      .toEqual([...beforeParts.keys()]);
    for (const [uri, bytes] of beforeParts) {
      expect(document.opcPackage.requirePart(uri).bytes, uri).toEqual(bytes);
    }
  });

  it('edits table-cell text directions through duplicate, rollback, and reopen lifecycles', async () => {
    const document = await PptxDocument.open(await tableTextDirectionFixture());
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
    const slide = document.slides[0]!;
    const table = slide.shapes[0] as TableModel;
    expect(table).toBeInstanceOf(TableModel);
    expect(table.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'horz',
      'vert',
      'vert270',
      'wordArtVert',
      undefined,
    ]);
    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;

    table.setCellTextDirection(0, 0, 'wordArtVert');
    table.setCellTextDirection(0, 1, 'horz');
    table.setCellTextDirection(0, 2, 'vert');
    table.setCellTextDirection(0, 3, 'vert270');
    table.setCellTextDirection(0, 4, 'wordArtVert');
    table.setCellTextDirection(1, 0, undefined);
    table.setCellText(1, 1, 'Edited text');
    table.setCellTextDirection(1, 2, 'vert');
    table.setTransform({ x: inches(2) });
    const snapshot = table.rows;
    (snapshot[0]!.cells[0] as { textDirection?: string }).textDirection = 'horz';

    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'wordArtVert',
      'horz',
      'vert',
      'vert270',
      'wordArtVert',
    ]);
    expect(table.rows[1]!.cells.map(({ text, textDirection }) => [text, textDirection])).toEqual([
      ['Clear me', undefined],
      ['Edited text', undefined],
      ['Merged placeholder', 'vert'],
      ['Neighbor', 'horz'],
      ['Tail', undefined],
    ]);
    const editedXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(editedXml).toContain('<x:keep xmlns:x="urn:test">OPAQUE</x:keep>');
    expect(editedXml).toContain('<a:tc hMerge="1">');
    expect(editedXml).toContain('<a:tcPr vert="horz" keep="NEIGHBOR"/>');
    expect(editedXml).toContain('<a:off x="1828800" y="914400"/>');
    expect(duplicateTable.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'horz',
      'vert',
      'vert270',
      'wordArtVert',
      undefined,
    ]);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    const rollbackJournal = [...document.opcPackage.mutations];
    expect(() =>
      document.transaction(() => {
        table.setCellTextDirection(0, 0, 'vert');
        table.setCellTextDirection(0, 1, undefined);
        throw new Error('restore public table edits');
      }),
    ).toThrow('restore public table edits');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[0]!.cells[0]!.textDirection).toBe('wordArtVert');

    const reopened = await PptxDocument.open(await document.write());
    const reopenedEdited = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(reopenedEdited.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'wordArtVert',
      'horz',
      'vert',
      'vert270',
      'wordArtVert',
    ]);
    expect(reopenedDuplicate.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'horz',
      'vert',
      'vert270',
      'wordArtVert',
      undefined,
    ]);
  });

  it('rejects invalid table-cell directions and physical coordinates before mutation', async () => {
    const document = await PptxDocument.open(await tableTextDirectionFixture());
    const slide = document.slides[0]!;
    const table = slide.shapes[0] as TableModel;
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    const directions = table.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection));
    const text = table.rows.map(({ cells }) => cells.map((cell) => cell.text));

    const invalidValues = [
      null,
      false,
      true,
      0,
      1,
      '',
      'Vert',
      ' vert ',
      'eaVert',
      'mongolianVert',
      'wordArtVertRtl',
      {},
      [],
      Symbol('direction'),
    ];
    for (const value of invalidValues) {
      expect(() => table.setCellTextDirection(0, 0, value as never)).toThrow(TypeError);
    }
    const invalidCoordinates = [
      [-1, 0],
      [0, -1],
      [0.5, 0],
      [0, 0.5],
      [Number.NaN, 0],
      [0, Number.NaN],
      [Number.POSITIVE_INFINITY, 0],
      [0, Number.NEGATIVE_INFINITY],
      [3, 0],
      [0, 5],
    ];
    for (const [row, column] of invalidCoordinates) {
      expect(() => table.setCellTextDirection(row!, column!, 'vert')).toThrow(RangeError);
    }

    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(document.slides).toHaveLength(1);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection))).toEqual(directions);
    expect(table.rows.map(({ cells }) => cells.map((cell) => cell.text))).toEqual(text);
  });

  it('edits table-cell horizontal alignments through duplicate, rollback, and reopen lifecycles', async () => {
    const document = await PptxDocument.open(await tableHorizontalAlignmentFixture());
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
    const slide = document.slides[0]!;
    const table = slide.shapes[0] as TableModel;
    const originalAlignments = ['left', 'center', 'right', 'justify', undefined];
    expect(table.rows[2]!.cells.map(({ horizontalAlignment }) => horizontalAlignment))
      .toEqual(originalAlignments);
    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;
    const duplicateBefore = document.opcPackage.requirePart(duplicate.partUri).bytes.slice();

    table.setCellHorizontalAlignment(2, 0, 'center');
    table.setCellHorizontalAlignment(2, 1, 'right');
    table.setCellHorizontalAlignment(2, 2, undefined);
    table.setCellHorizontalAlignment(2, 3, 'left');
    table.setCellHorizontalAlignment(2, 4, 'justify');
    table.setCellText(2, 4, 'Edited merged alignment');
    table.setCellTextFit(2, 0, 'resize');
    table.setCellTextDirection(2, 3, 'horz');
    table.setCellVerticalAlignment(2, 1, 'top');
    table.setCellMargins(2, 2, { left: 9, right: 10 });
    table.setCellBorders(2, 0, { kind: 'none' });
    table.setCellFill(2, 3, { kind: 'none' });
    table.setTransform({ x: inches(2) });
    table.setColumnWidths([
      inches(1),
      inches(1.5),
      inches(2),
      inches(1.5),
      inches(2),
    ]);
    table.setRowHeights([inches(1), inches(1.25), inches(0.75)]);
    const snapshot = table.rows;
    (snapshot[2]!.cells[1] as { horizontalAlignment?: string }).horizontalAlignment = 'left';

    const editedAlignments = ['center', 'right', undefined, 'left', 'justify'];
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[2]!.cells.map(({ horizontalAlignment }) => horizontalAlignment))
      .toEqual(editedAlignments);
    expect(table.rows[2]!.cells.map(({ textFit }) => textFit)).toEqual([
      'resize',
      'shrink',
      'resize',
      undefined,
      undefined,
    ]);
    expect(table.rows[2]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'horz',
      'vert',
      'vert270',
      'horz',
      'wordArtVert',
    ]);
    expect(table.rows[2]!.cells.map(({ verticalAlignment }) => verticalAlignment)).toEqual([
      'top',
      'top',
      'bottom',
      undefined,
      'middle',
    ]);
    expect(table.rows[2]!.cells[2]!.margins).toEqual({ left: 9, right: 10 });
    expect(table.rows[2]!.cells[0]!.borders).toEqual({
      top: { kind: 'none' },
      right: { kind: 'none' },
      bottom: { kind: 'none' },
      left: { kind: 'none' },
    });
    expect(table.rows[2]!.cells[3]!.fill).toEqual({ kind: 'none' });
    expect(table.rows[2]!.cells[4]!.text).toBe('Edited merged alignment');
    expect(table.columnWidths).toEqual([
      inches(1),
      inches(1.5),
      inches(2),
      inches(1.5),
      inches(2),
    ]);
    expect(table.rowHeights).toEqual([inches(1), inches(1.25), inches(0.75)]);

    const editedXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(editedXml).toContain('<a:pPr algn="ctr" keep="LEFT"/>');
    expect(editedXml).toContain('<a:pPr algn="r" keep="CENTER"><a:buNone/></a:pPr>');
    expect(editedXml).toContain('<a:pPr keep="RIGHT"/>');
    expect(editedXml).toContain('<a:pPr algn="l" keep="JUSTIFY"/>');
    expect(editedXml).toContain('<a:pPr keep="ABSENT" algn="just"/>');
    expect(editedXml).toContain('<a:bodyPr custom="NONE"><a:spAutoFit/></a:bodyPr>');
    expect(editedXml).toContain('keep="FIT-SHRINK"');
    expect(editedXml).toContain('keep="FIT-RESIZE"');
    expect(editedXml).toContain('keep="FIT-ABSENT"');
    expect(editedXml).toContain('keep="FIT-MERGED"');
    expect(editedXml).not.toMatch(/<a:tcPr[^>]*\salgn=/);
    expect(document.opcPackage.requirePart(duplicate.partUri).bytes).toEqual(duplicateBefore);
    expect(duplicateTable.rows[2]!.cells.map(({ horizontalAlignment }) => horizontalAlignment))
      .toEqual(originalAlignments);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const rollbackJournal = [...document.opcPackage.mutations];
    expect(() =>
      document.transaction(() => {
        table.setCellHorizontalAlignment(2, 0, 'left');
        table.setCellHorizontalAlignment(2, 4, undefined);
        throw new Error('restore public table horizontal alignment edits');
      }),
    ).toThrow('restore public table horizontal alignment edits');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[2]!.cells.map(({ horizontalAlignment }) => horizontalAlignment))
      .toEqual(editedAlignments);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedEdited = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(reopenedEdited.rows[2]!.cells.map(({ horizontalAlignment }) => horizontalAlignment))
      .toEqual(editedAlignments);
    expect(reopenedEdited.rows[2]!.cells[4]!.text).toBe('Edited merged alignment');
    expect(reopenedEdited.columnWidths).toEqual(table.columnWidths);
    expect(reopenedEdited.rowHeights).toEqual(table.rowHeights);
    expect(reopenedDuplicate.rows[2]!.cells.map(({ horizontalAlignment }) => horizontalAlignment))
      .toEqual(originalAlignments);
  });

  it('rejects invalid table-cell horizontal alignments and physical coordinates before mutation', async () => {
    const document = await PptxDocument.open(await tableHorizontalAlignmentFixture());
    const slide = document.slides[0]!;
    const table = slide.shapes[0] as TableModel;
    const before = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const journal = [...document.opcPackage.mutations];
    const snapshots = table.rows.map(({ cells }) => cells.map((cell) => ({
      text: cell.text,
      horizontalAlignment: cell.horizontalAlignment,
      verticalAlignment: cell.verticalAlignment,
      textDirection: cell.textDirection,
      textFit: cell.textFit,
      margins: cell.margins,
      borders: cell.borders,
      fill: cell.fill,
    })));

    const invalidValues = [
      null,
      false,
      true,
      0,
      1,
      '',
      'Left',
      ' center ',
      'l',
      'ctr',
      'r',
      'just',
      'dist',
      'thaiDist',
      'justLow',
      {},
      [],
      Symbol('horizontal alignment'),
    ];
    for (const value of invalidValues) {
      expect(() => table.setCellHorizontalAlignment(2, 0, value as never)).toThrow(TypeError);
    }
    const invalidCoordinates = [
      [-1, 0],
      [0, -1],
      [0.5, 0],
      [0, 0.5],
      [Number.NaN, 0],
      [0, Number.NaN],
      [Number.POSITIVE_INFINITY, 0],
      [0, Number.NEGATIVE_INFINITY],
      [3, 0],
      [2, 5],
    ];
    for (const [row, column] of invalidCoordinates) {
      expect(() => table.setCellHorizontalAlignment(row!, column!, 'center')).toThrow(RangeError);
    }

    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(document.slides).toHaveLength(1);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows.map(({ cells }) => cells.map((cell) => ({
      text: cell.text,
      horizontalAlignment: cell.horizontalAlignment,
      verticalAlignment: cell.verticalAlignment,
      textDirection: cell.textDirection,
      textFit: cell.textFit,
      margins: cell.margins,
      borders: cell.borders,
      fill: cell.fill,
    })))).toEqual(snapshots);
  });

  it('edits table-cell fits through duplicate, rollback, and reopen lifecycles', async () => {
    const document = await PptxDocument.open(await tableTextDirectionFixture());
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
    const slide = document.slides[0]!;
    const table = slide.shapes[0] as TableModel;
    expect(table.rows[2]!.cells.map(({ textFit }) => textFit)).toEqual([
      'none',
      'shrink',
      'resize',
      undefined,
      undefined,
    ]);
    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;

    table.setCellTextFit(2, 0, 'none');
    table.setCellTextFit(2, 1, 'resize');
    table.setCellTextFit(2, 2, 'shrink');
    table.setCellTextFit(2, 3, 'resize');
    table.setCellTextFit(2, 4, 'shrink');
    table.setCellText(2, 4, 'Edited merged fit');
    table.setCellTextDirection(2, 3, 'horz');
    table.setTransform({ x: inches(2) });
    const snapshot = table.rows;
    (snapshot[2]!.cells[1] as { textFit?: string }).textFit = 'none';

    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[2]!.cells.map(({ textFit }) => textFit)).toEqual([
      undefined,
      'resize',
      'shrink',
      'resize',
      'shrink',
    ]);
    expect(table.rows[2]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'horz',
      'vert',
      'vert270',
      'horz',
      'wordArtVert',
    ]);
    expect(table.rows[2]!.cells[4]!.text).toBe('Edited merged fit');
    const editedXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(editedXml).toContain('<a:bodyPr custom="NONE"></a:bodyPr>');
    expect(editedXml).toContain('<a:bodyPr custom="SHRINK"><a:spAutoFit/></a:bodyPr>');
    expect(editedXml).toContain('<a:bodyPr custom="RESIZE"><a:normAutofit/></a:bodyPr>');
    expect(editedXml).toContain('<a:bodyPr><a:spAutoFit/></a:bodyPr>');
    expect(editedXml).toContain('<a:tc hMerge="1">');
    expect(editedXml).toContain('<a:tcPr vert="wordArtVert" anchor="ctr" marT="45720" marR="91440" marB="45720" marL="91440" keep="FIT-MERGED"><a:solidFill><a:srgbClr val="70AD47"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:tcPr>');
    expect(editedXml).toContain('<a:off x="1828800" y="914400"/>');
    expect(duplicateTable.rows[2]!.cells.map(({ textFit }) => textFit)).toEqual([
      'none',
      'shrink',
      'resize',
      undefined,
      undefined,
    ]);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    const rollbackJournal = [...document.opcPackage.mutations];
    expect(() =>
      document.transaction(() => {
        table.setCellTextFit(2, 1, 'shrink');
        table.setCellTextFit(2, 3, undefined);
        throw new Error('restore public table fit edits');
      }),
    ).toThrow('restore public table fit edits');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[2]!.cells.map(({ textFit }) => textFit)).toEqual([
      undefined,
      'resize',
      'shrink',
      'resize',
      'shrink',
    ]);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedEdited = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(reopenedEdited.rows[2]!.cells.map(({ textFit }) => textFit)).toEqual([
      undefined,
      'resize',
      'shrink',
      'resize',
      'shrink',
    ]);
    expect(reopenedDuplicate.rows[2]!.cells.map(({ textFit }) => textFit)).toEqual([
      'none',
      'shrink',
      'resize',
      undefined,
      undefined,
    ]);
  });

  it('rejects invalid table-cell fits and physical coordinates before mutation', async () => {
    const document = await PptxDocument.open(await tableTextDirectionFixture());
    const slide = document.slides[0]!;
    const table = slide.shapes[0] as TableModel;
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    const fits = table.rows.map(({ cells }) => cells.map(({ textFit }) => textFit));
    const directions = table.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection));
    const text = table.rows.map(({ cells }) => cells.map((cell) => cell.text));

    const invalidValues = [
      null,
      false,
      true,
      0,
      1,
      '',
      'None',
      ' shrink ',
      'grow',
      {},
      [],
      Symbol('fit'),
    ];
    for (const value of invalidValues) {
      expect(() => table.setCellTextFit(2, 0, value as never)).toThrow(TypeError);
    }
    const invalidCoordinates = [
      [-1, 0],
      [0, -1],
      [0.5, 0],
      [0, 0.5],
      [Number.NaN, 0],
      [0, Number.NaN],
      [Number.POSITIVE_INFINITY, 0],
      [0, Number.NEGATIVE_INFINITY],
      [3, 0],
      [2, 5],
    ];
    for (const [row, column] of invalidCoordinates) {
      expect(() => table.setCellTextFit(row!, column!, 'shrink')).toThrow(RangeError);
    }

    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(document.slides).toHaveLength(1);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows.map(({ cells }) => cells.map(({ textFit }) => textFit))).toEqual(fits);
    expect(table.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection))).toEqual(directions);
    expect(table.rows.map(({ cells }) => cells.map((cell) => cell.text))).toEqual(text);
  });

  it('edits table-cell vertical alignments through duplicate, rollback, and reopen lifecycles', async () => {
    const document = await PptxDocument.open(await tableTextDirectionFixture());
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
    const slide = document.slides[0]!;
    const table = slide.shapes[0] as TableModel;
    expect(table.rows[2]!.cells.map(({ verticalAlignment }) => verticalAlignment)).toEqual([
      'top',
      'middle',
      'bottom',
      undefined,
      'middle',
    ]);
    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;

    table.setCellVerticalAlignment(2, 0, 'middle');
    table.setCellVerticalAlignment(2, 1, 'top');
    table.setCellVerticalAlignment(2, 2, undefined);
    table.setCellVerticalAlignment(2, 3, 'bottom');
    table.setCellVerticalAlignment(2, 4, 'top');
    table.setCellTextFit(2, 0, 'resize');
    table.setCellTextDirection(2, 3, 'horz');
    table.setCellText(2, 4, 'Edited merged alignment');
    table.setTransform({ x: inches(2) });
    const snapshot = table.rows;
    (snapshot[2]!.cells[1] as { verticalAlignment?: string }).verticalAlignment = 'bottom';

    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[2]!.cells.map(({ verticalAlignment }) => verticalAlignment)).toEqual([
      'middle',
      'top',
      undefined,
      'bottom',
      'top',
    ]);
    expect(table.rows[2]!.cells.map(({ textFit }) => textFit)).toEqual([
      'resize',
      'shrink',
      'resize',
      undefined,
      undefined,
    ]);
    expect(table.rows[2]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'horz',
      'vert',
      'vert270',
      'horz',
      'wordArtVert',
    ]);
    expect(table.rows[2]!.cells[4]!.text).toBe('Edited merged alignment');
    const editedXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(editedXml).toContain('<a:tcPr vert="horz" anchor="ctr" marT="50800" marR="101600" marB="152400" marL="203200" keep="FIT-NONE"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:tcPr>');
    expect(editedXml).toContain('<a:tcPr vert="vert" anchor="t" marT="0" marL="25400" keep="FIT-SHRINK"><a:solidFill><a:schemeClr val="accent2"><a:alpha val="75000"/></a:schemeClr></a:solidFill></a:tcPr>');
    expect(editedXml).toContain('<a:tcPr vert="vert270" marR="91440" keep="FIT-RESIZE"><a:noFill/></a:tcPr>');
    expect(editedXml).toContain('<a:tcPr keep="FIT-ABSENT" anchor="b" vert="horz"/>');
    expect(editedXml).toContain('<a:tcPr vert="wordArtVert" anchor="t" marT="45720" marR="91440" marB="45720" marL="91440" keep="FIT-MERGED"><a:solidFill><a:srgbClr val="70AD47"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:tcPr>');
    expect(editedXml).toContain('<a:bodyPr custom="NONE"><a:spAutoFit/></a:bodyPr>');
    expect(editedXml).toContain('<a:off x="1828800" y="914400"/>');
    expect(duplicateTable.rows[2]!.cells.map(({ verticalAlignment }) => verticalAlignment)).toEqual([
      'top',
      'middle',
      'bottom',
      undefined,
      'middle',
    ]);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    const rollbackJournal = [...document.opcPackage.mutations];
    expect(() =>
      document.transaction(() => {
        table.setCellVerticalAlignment(2, 1, 'bottom');
        table.setCellVerticalAlignment(2, 3, undefined);
        throw new Error('restore public table vertical alignment edits');
      }),
    ).toThrow('restore public table vertical alignment edits');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[2]!.cells.map(({ verticalAlignment }) => verticalAlignment)).toEqual([
      'middle',
      'top',
      undefined,
      'bottom',
      'top',
    ]);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedEdited = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(reopenedEdited.rows[2]!.cells.map(({ verticalAlignment }) => verticalAlignment)).toEqual([
      'middle',
      'top',
      undefined,
      'bottom',
      'top',
    ]);
    expect(reopenedDuplicate.rows[2]!.cells.map(({ verticalAlignment }) => verticalAlignment)).toEqual([
      'top',
      'middle',
      'bottom',
      undefined,
      'middle',
    ]);
  });

  it('projects and edits table-level borders through the public root API', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const table = slide.addTable([
      ['North', 'South'],
      ['East', 'West'],
    ], {
      name: 'Public table borders',
      border: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent1' },
        width: 1.5,
        style: 'dash',
      },
      columnWidths: inches(2),
      rowHeights: inches(0.75),
    });
    const four = <T>(border: T) => ({
      top: border,
      right: border,
      bottom: border,
      left: border,
    });
    const initialLine = {
      kind: 'line' as const,
      color: { kind: 'scheme' as const, value: 'accent1' as const },
      width: 1.5,
      style: 'dash' as const,
    };
    const initialBorders = four(initialLine);
    expect(table).toBeInstanceOf(TableModel);
    expect(table.borders).toEqual(initialBorders);
    expect(validatePackage(document.opcPackage)
      .filter(({ severity }) => severity === 'error')).toEqual([]);

    const detached = table.borders!;
    const detachedTop = detached.top;
    expect(detachedTop?.kind).toBe('line');
    if (detachedTop?.kind === 'line') {
      (detachedTop.color as { value: string }).value = 'accent6';
    }
    expect(table.borders).toEqual(initialBorders);
    const noOpBytes = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...document.opcPackage.mutations];
    const noOpDiagnostics = [...document.diagnostics];
    table.borders = initialLine;
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(noOpBytes);
    expect(document.opcPackage.mutations).toEqual(noOpJournal);
    expect(document.diagnostics).toEqual(noOpDiagnostics);

    table.setCellBorders(0, 1, { kind: 'none' });
    expect(table.borders).toBeUndefined();
    const partial = {
      top: {
        kind: 'line' as const,
        color: { kind: 'srgb' as const, value: 'D9EAF7' },
        width: 2,
      },
      bottom: { kind: 'none' as const },
    };
    table.borders = partial;
    expect(table.borders).toEqual(partial);
    expect(table.rows.flatMap(({ cells }) => cells).map(({ borders }) => borders))
      .toEqual(Array(4).fill(partial));

    table.borders = { kind: 'none' };
    const noneBorders = four({ kind: 'none' as const });
    expect(table.borders).toEqual(noneBorders);
    expect(table.rows.flatMap(({ cells }) => cells).map(({ borders }) => borders))
      .toEqual(Array(4).fill(noneBorders));

    table.borders = undefined;
    expect(table.borders).toBeUndefined();
    expect(table.rows.flatMap(({ cells }) => cells).map(({ borders }) => borders))
      .toEqual([undefined, undefined, undefined, undefined]);
    const clearBytes = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const clearJournal = [...document.opcPackage.mutations];
    table.borders = {};
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(clearBytes);
    expect(document.opcPackage.mutations).toEqual(clearJournal);

    table.borders = { kind: 'none' };
    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;
    expect(duplicateTable.borders).toEqual(noneBorders);
    const finalLine = {
      kind: 'line' as const,
      color: { kind: 'scheme' as const, value: 'accent2' as const },
      width: 0,
      style: 'solid' as const,
    };
    const finalBorders = four(finalLine);
    table.borders = finalLine;
    expect(table.borders).toEqual(finalBorders);
    expect(duplicateTable.borders).toEqual(noneBorders);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const rollbackJournal = [...document.opcPackage.mutations];
    expect(() => document.transaction(() => {
      table.borders = { kind: 'none' };
      throw new Error('restore public table borders');
    })).toThrow('restore public table borders');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(table.borders).toEqual(finalBorders);

    const beforeInvalid = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const invalidJournal = [...document.opcPackage.mutations];
    const invalidDiagnostics = [...document.diagnostics];
    expect(() => {
      table.borders = null as never;
    }).toThrow('Table borders must be an object');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeInvalid);
    expect(document.opcPackage.mutations).toEqual(invalidJournal);
    expect(document.diagnostics).toEqual(invalidDiagnostics);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(reopenedTable.borders).toEqual(finalBorders);
    expect(reopenedTable.rows.flatMap(({ cells }) => cells).map(({ borders }) => borders))
      .toEqual(Array(4).fill(finalBorders));
    expect(reopenedDuplicate.borders).toEqual(noneBorders);
    expect(reopenedDuplicate.rows.flatMap(({ cells }) => cells).map(({ borders }) => borders))
      .toEqual(Array(4).fill(noneBorders));
    expect(validatePackage(reopened.opcPackage)
      .filter(({ severity }) => severity === 'error')).toEqual([]);

    if (false) {
      const borders: TableCellBorders | undefined = table.borders;
      table.borders = { kind: 'none' };
      table.borders = [
        { kind: 'none' },
        undefined,
        { kind: 'line', color: { kind: 'scheme', value: 'accent1' }, width: 2 },
        undefined,
      ];
      table.borders = undefined;
      // @ts-expect-error table borders reject null
      table.borders = null;
      // @ts-expect-error table borders reject PptxGenJS-shaped input
      table.borders = { type: 'solid', color: '4472C4', pt: 1 };
      void borders;
    }
  });

  it('projects and edits table-level fill through the public root API', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const table = slide.addTable([
      ['North', 'South'],
      ['East', 'West'],
    ], {
      name: 'Public table fill',
      fill: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 25,
      },
      columnWidths: inches(2),
      rowHeights: inches(0.75),
    });
    const initialFill = {
      kind: 'solid' as const,
      color: { kind: 'scheme' as const, value: 'accent1' as const },
      transparency: 25,
    };
    expect(table).toBeInstanceOf(TableModel);
    expect(table.fill).toEqual(initialFill);
    expect(validatePackage(document.opcPackage)
      .filter(({ severity }) => severity === 'error')).toEqual([]);

    const detached = table.fill!;
    expect(detached.kind).toBe('solid');
    if (detached.kind === 'solid') {
      (detached.color as { value: string }).value = 'accent6';
    }
    expect(table.fill).toEqual(initialFill);
    const noOpBytes = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...document.opcPackage.mutations];
    const noOpDiagnostics = [...document.diagnostics];
    table.fill = initialFill;
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(noOpBytes);
    expect(document.opcPackage.mutations).toEqual(noOpJournal);
    expect(document.diagnostics).toEqual(noOpDiagnostics);

    table.setCellFill(0, 1, { kind: 'none' });
    expect(table.fill).toBeUndefined();
    table.fill = { kind: 'none' };
    expect(table.fill).toEqual({ kind: 'none' });
    expect(table.rows.flatMap(({ cells }) => cells).map(({ fill }) => fill))
      .toEqual(Array(4).fill({ kind: 'none' }));

    table.fill = {
      kind: 'solid',
      color: { kind: 'srgb', value: 'D9EAF7' },
      transparency: 0,
    };
    const explicitZero = {
      kind: 'solid' as const,
      color: { kind: 'srgb' as const, value: 'D9EAF7' },
      transparency: 0,
    };
    expect(table.fill).toEqual(explicitZero);
    expect(table.rows.flatMap(({ cells }) => cells).map(({ fill }) => fill))
      .toEqual(Array(4).fill(explicitZero));

    table.fill = undefined;
    expect(table.fill).toBeUndefined();
    expect(table.rows.flatMap(({ cells }) => cells).map(({ fill }) => fill))
      .toEqual([undefined, undefined, undefined, undefined]);
    const clearBytes = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const clearJournal = [...document.opcPackage.mutations];
    table.fill = undefined;
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(clearBytes);
    expect(document.opcPackage.mutations).toEqual(clearJournal);

    table.fill = { kind: 'none' };
    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;
    expect(duplicateTable.fill).toEqual({ kind: 'none' });
    table.fill = {
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 50,
    };
    const finalFill = {
      kind: 'solid' as const,
      color: { kind: 'scheme' as const, value: 'accent2' as const },
      transparency: 50,
    };
    expect(table.fill).toEqual(finalFill);
    expect(duplicateTable.fill).toEqual({ kind: 'none' });

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const rollbackJournal = [...document.opcPackage.mutations];
    expect(() => document.transaction(() => {
      table.fill = { kind: 'none' };
      throw new Error('restore public table fill');
    })).toThrow('restore public table fill');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(table.fill).toEqual(finalFill);

    const beforeInvalid = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const invalidJournal = [...document.opcPackage.mutations];
    const invalidDiagnostics = [...document.diagnostics];
    expect(() => {
      table.fill = null as never;
    }).toThrow('Table fill must be an object');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeInvalid);
    expect(document.opcPackage.mutations).toEqual(invalidJournal);
    expect(document.diagnostics).toEqual(invalidDiagnostics);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(reopenedTable.fill).toEqual(finalFill);
    expect(reopenedTable.rows.flatMap(({ cells }) => cells).map(({ fill }) => fill))
      .toEqual(Array(4).fill(finalFill));
    expect(reopenedDuplicate.fill).toEqual({ kind: 'none' });
    expect(reopenedDuplicate.rows.flatMap(({ cells }) => cells).map(({ fill }) => fill))
      .toEqual(Array(4).fill({ kind: 'none' }));
    expect(validatePackage(reopened.opcPackage)
      .filter(({ severity }) => severity === 'error')).toEqual([]);

    if (false) {
      const fill: TableCellFill | undefined = table.fill;
      table.fill = { kind: 'none' };
      table.fill = { kind: 'solid', color: { kind: 'srgb', value: 'D9EAF7' } };
      table.fill = undefined;
      // @ts-expect-error table fill rejects null
      table.fill = null;
      // @ts-expect-error table fill rejects PptxGenJS-shaped input
      table.fill = { color: 'D9EAF7' };
      void fill;
    }
  });

  it('projects and edits table-level margins through the public root API', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const table = slide.addTable([
      ['North', 'South'],
      ['East', 'West'],
    ], {
      name: 'Public table margins',
      margin: [3.6, 7.2, 10.8, 14.4],
      columnWidths: inches(2),
      rowHeights: inches(0.75),
    });
    expect(table).toBeInstanceOf(TableModel);
    expect(table.margins).toEqual({ top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 });
    expect(validatePackage(document.opcPackage)
      .filter(({ severity }) => severity === 'error')).toEqual([]);

    const detached = table.margins!;
    (detached as { top?: number }).top = 99;
    expect(table.margins).toEqual({ top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 });
    const noOpBytes = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...document.opcPackage.mutations];
    const noOpDiagnostics = [...document.diagnostics];
    table.margins = [3.6, 7.2, 10.8, 14.4];
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(noOpBytes);
    expect(document.opcPackage.mutations).toEqual(noOpJournal);
    expect(document.diagnostics).toEqual(noOpDiagnostics);

    table.setCellMargins(0, 1, { top: 9 });
    expect(table.margins).toBeUndefined();
    table.margins = 6;
    expect(table.margins).toEqual({ top: 6, right: 6, bottom: 6, left: 6 });
    expect(table.rows.flatMap(({ cells }) => cells).map(({ margins }) => margins))
      .toEqual(Array(4).fill({ top: 6, right: 6, bottom: 6, left: 6 }));

    table.margins = { top: 2, left: 4 };
    expect(table.margins).toEqual({ top: 2, left: 4 });
    expect(table.rows.flatMap(({ cells }) => cells).map(({ margins }) => margins))
      .toEqual(Array(4).fill({ top: 2, left: 4 }));
    table.margins = {};
    expect(table.margins).toBeUndefined();
    expect(table.rows.flatMap(({ cells }) => cells)
      .every(({ margins }) => margins === undefined)).toBe(true);
    const clearedBytes = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const clearedJournal = [...document.opcPackage.mutations];
    table.margins = undefined;
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(clearedBytes);
    expect(document.opcPackage.mutations).toEqual(clearedJournal);

    table.margins = [1, 2, 3, 4];
    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;
    expect(duplicateTable.margins).toEqual({ top: 1, right: 2, bottom: 3, left: 4 });
    table.margins = 5;
    expect(table.margins).toEqual({ top: 5, right: 5, bottom: 5, left: 5 });
    expect(duplicateTable.margins).toEqual({ top: 1, right: 2, bottom: 3, left: 4 });
    table.margins = [1, 2, 3, 4];

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const rollbackJournal = [...document.opcPackage.mutations];
    expect(() => document.transaction(() => {
      table.margins = { bottom: 9 };
      throw new Error('restore public table margins');
    })).toThrow('restore public table margins');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(table.margins).toEqual({ top: 1, right: 2, bottom: 3, left: 4 });

    const beforeInvalid = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const invalidJournal = [...document.opcPackage.mutations];
    const invalidDiagnostics = [...document.diagnostics];
    expect(() => {
      table.margins = null as never;
    }).toThrow('Table margins must be a number, four-value tuple, or margin object');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeInvalid);
    expect(document.opcPackage.mutations).toEqual(invalidJournal);
    expect(document.diagnostics).toEqual(invalidDiagnostics);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    for (const current of [reopenedTable, reopenedDuplicate]) {
      expect(current.margins).toEqual({ top: 1, right: 2, bottom: 3, left: 4 });
      expect(current.rows.flatMap(({ cells }) => cells).map(({ margins }) => margins))
        .toEqual(Array(4).fill({ top: 1, right: 2, bottom: 3, left: 4 }));
    }
    expect(validatePackage(reopened.opcPackage)
      .filter(({ severity }) => severity === 'error')).toEqual([]);

    if (false) {
      const margins: TextBoxMargins | undefined = table.margins;
      table.margins = 6;
      table.margins = [1, 2, 3, 4];
      table.margins = { top: 2, left: 4 };
      table.margins = {};
      table.margins = undefined;
      // @ts-expect-error table margins reject null
      table.margins = null;
      // @ts-expect-error table margin tuple requires four values
      table.margins = [1, 2, 3];
      void margins;
    }
  });

  it('projects and edits table-level horizontal alignment through the public root API', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const table = slide.addTable([
      ['North', 'South'],
      ['East', 'West'],
    ], {
      name: 'Public table horizontal alignment',
      align: 'center',
      columnWidths: inches(2),
      rowHeights: inches(0.75),
    });
    expect(table).toBeInstanceOf(TableModel);
    expect(table.horizontalAlignment).toBe('center');
    expect(validatePackage(document.opcPackage)
      .filter(({ severity }) => severity === 'error')).toEqual([]);

    const noOpBytes = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...document.opcPackage.mutations];
    const noOpDiagnostics = [...document.diagnostics];
    void table.horizontalAlignment;
    table.horizontalAlignment = 'center';
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(noOpBytes);
    expect(document.opcPackage.mutations).toEqual(noOpJournal);
    expect(document.diagnostics).toEqual(noOpDiagnostics);

    table.setCellHorizontalAlignment(0, 1, 'right');
    expect(table.horizontalAlignment).toBeUndefined();
    table.horizontalAlignment = 'justify';
    expect(table.horizontalAlignment).toBe('justify');
    expect(table.rows.flatMap(({ cells }) => cells)
      .map(({ horizontalAlignment }) => horizontalAlignment))
      .toEqual(['justify', 'justify', 'justify', 'justify']);

    table.horizontalAlignment = 'left';
    expect(table.horizontalAlignment).toBe('left');
    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;
    expect(duplicateTable.horizontalAlignment).toBe('left');
    table.horizontalAlignment = undefined;
    expect(table.horizontalAlignment).toBeUndefined();
    expect(table.rows.flatMap(({ cells }) => cells)
      .every(({ horizontalAlignment }) => horizontalAlignment === undefined)).toBe(true);
    expect(duplicateTable.horizontalAlignment).toBe('left');
    table.horizontalAlignment = 'right';

    const beforeInvalid = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const invalidJournal = [...document.opcPackage.mutations];
    const invalidDiagnostics = [...document.diagnostics];
    expect(() => {
      table.horizontalAlignment = 'dist' as never;
    }).toThrow('Table horizontal alignment must be left, center, right, or justify');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeInvalid);
    expect(document.opcPackage.mutations).toEqual(invalidJournal);
    expect(document.diagnostics).toEqual(invalidDiagnostics);
    expect(table.horizontalAlignment).toBe('right');

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(reopenedTable.horizontalAlignment).toBe('right');
    expect(reopenedDuplicate.horizontalAlignment).toBe('left');
    expect(reopenedTable.rows.flatMap(({ cells }) => cells)
      .map(({ horizontalAlignment }) => horizontalAlignment))
      .toEqual(['right', 'right', 'right', 'right']);
    expect(reopenedDuplicate.rows.flatMap(({ cells }) => cells)
      .map(({ horizontalAlignment }) => horizontalAlignment))
      .toEqual(['left', 'left', 'left', 'left']);
    expect(validatePackage(reopened.opcPackage)
      .filter(({ severity }) => severity === 'error')).toEqual([]);

    if (false) {
      const alignment: TextAlignment | undefined = table.horizontalAlignment;
      table.horizontalAlignment = 'left';
      table.horizontalAlignment = 'center';
      table.horizontalAlignment = 'right';
      table.horizontalAlignment = 'justify';
      table.horizontalAlignment = undefined;
      // @ts-expect-error unsupported table-level horizontal alignment
      table.horizontalAlignment = 'dist';
      void alignment;
    }
  });

  it('projects and edits table-level vertical alignment through the public root API', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const table = slide.addTable([
      ['North', 'South'],
      ['East', 'West'],
    ], {
      name: 'Public table vertical alignment',
      valign: 'middle',
      columnWidths: inches(2),
      rowHeights: inches(0.75),
    });
    expect(table).toBeInstanceOf(TableModel);
    expect(table.verticalAlignment).toBe('middle');
    expect(validatePackage(document.opcPackage)
      .filter(({ severity }) => severity === 'error')).toEqual([]);

    const noOpBytes = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...document.opcPackage.mutations];
    const noOpDiagnostics = [...document.diagnostics];
    void table.verticalAlignment;
    table.verticalAlignment = 'middle';
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(noOpBytes);
    expect(document.opcPackage.mutations).toEqual(noOpJournal);
    expect(document.diagnostics).toEqual(noOpDiagnostics);

    table.setCellVerticalAlignment(0, 1, 'top');
    expect(table.verticalAlignment).toBeUndefined();
    table.verticalAlignment = 'top';
    expect(table.verticalAlignment).toBe('top');
    expect(table.rows.flatMap(({ cells }) => cells)
      .map(({ verticalAlignment }) => verticalAlignment))
      .toEqual(['top', 'top', 'top', 'top']);

    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;
    expect(duplicateTable.verticalAlignment).toBe('top');
    table.verticalAlignment = undefined;
    expect(table.verticalAlignment).toBeUndefined();
    expect(table.rows.flatMap(({ cells }) => cells)
      .every(({ verticalAlignment }) => verticalAlignment === undefined)).toBe(true);
    expect(duplicateTable.verticalAlignment).toBe('top');
    table.verticalAlignment = 'bottom';

    const beforeInvalid = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const invalidJournal = [...document.opcPackage.mutations];
    const invalidDiagnostics = [...document.diagnostics];
    expect(() => {
      table.verticalAlignment = 'distributed' as never;
    }).toThrow('Table vertical alignment must be top, middle, or bottom');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeInvalid);
    expect(document.opcPackage.mutations).toEqual(invalidJournal);
    expect(document.diagnostics).toEqual(invalidDiagnostics);
    expect(table.verticalAlignment).toBe('bottom');

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(reopenedTable.verticalAlignment).toBe('bottom');
    expect(reopenedDuplicate.verticalAlignment).toBe('top');
    expect(reopenedTable.rows.flatMap(({ cells }) => cells)
      .map(({ verticalAlignment }) => verticalAlignment))
      .toEqual(['bottom', 'bottom', 'bottom', 'bottom']);
    expect(reopenedDuplicate.rows.flatMap(({ cells }) => cells)
      .map(({ verticalAlignment }) => verticalAlignment))
      .toEqual(['top', 'top', 'top', 'top']);
    expect(validatePackage(reopened.opcPackage)
      .filter(({ severity }) => severity === 'error')).toEqual([]);

    if (false) {
      const alignment: TextBoxVerticalAlignment | undefined = table.verticalAlignment;
      table.verticalAlignment = 'top';
      table.verticalAlignment = 'middle';
      table.verticalAlignment = 'bottom';
      table.verticalAlignment = undefined;
      // @ts-expect-error unsupported table-level vertical alignment
      table.verticalAlignment = 'distributed';
      void alignment;
    }
  });

  it('projects and edits table-level text direction through the public root API', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const table = slide.addTable([
      ['North', 'South'],
      ['East', 'West'],
    ], {
      name: 'Public table text direction',
      textDirection: 'vert270',
      columnWidths: inches(2),
      rowHeights: inches(0.75),
    });
    expect(table).toBeInstanceOf(TableModel);
    expect(table.textDirection).toBe('vert270');
    expect(validatePackage(document.opcPackage)
      .filter(({ severity }) => severity === 'error')).toEqual([]);

    const noOpBytes = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...document.opcPackage.mutations];
    const noOpDiagnostics = [...document.diagnostics];
    void table.textDirection;
    table.textDirection = 'vert270';
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(noOpBytes);
    expect(document.opcPackage.mutations).toEqual(noOpJournal);
    expect(document.diagnostics).toEqual(noOpDiagnostics);

    table.setCellTextDirection(0, 1, 'vert');
    expect(table.textDirection).toBeUndefined();
    table.textDirection = 'wordArtVert';
    expect(table.textDirection).toBe('wordArtVert');
    expect(table.rows.flatMap(({ cells }) => cells)
      .map(({ textDirection }) => textDirection))
      .toEqual(['wordArtVert', 'wordArtVert', 'wordArtVert', 'wordArtVert']);

    table.textDirection = 'horz';
    expect(table.textDirection).toBe('horz');
    expect(table.rows.flatMap(({ cells }) => cells)
      .every(({ textDirection }) => textDirection === 'horz')).toBe(true);
    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;
    expect(duplicateTable.textDirection).toBe('horz');

    table.textDirection = undefined;
    expect(table.textDirection).toBeUndefined();
    expect(table.rows.flatMap(({ cells }) => cells)
      .every(({ textDirection }) => textDirection === undefined)).toBe(true);
    expect(duplicateTable.textDirection).toBe('horz');
    table.textDirection = 'vert';

    const beforeInvalid = document.opcPackage.requirePart(slide.partUri).bytes.slice();
    const invalidJournal = [...document.opcPackage.mutations];
    const invalidDiagnostics = [...document.diagnostics];
    expect(() => {
      table.textDirection = 'eaVert' as never;
    }).toThrow('Table text direction must be horz, vert, vert270, or wordArtVert');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeInvalid);
    expect(document.opcPackage.mutations).toEqual(invalidJournal);
    expect(document.diagnostics).toEqual(invalidDiagnostics);
    expect(table.textDirection).toBe('vert');

    const reopened = await PptxDocument.open(await document.write());
    const reopenedTable = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(reopenedTable.textDirection).toBe('vert');
    expect(reopenedDuplicate.textDirection).toBe('horz');
    expect(reopenedTable.rows.flatMap(({ cells }) => cells)
      .map(({ textDirection }) => textDirection))
      .toEqual(['vert', 'vert', 'vert', 'vert']);
    expect(reopenedDuplicate.rows.flatMap(({ cells }) => cells)
      .map(({ textDirection }) => textDirection))
      .toEqual(['horz', 'horz', 'horz', 'horz']);
    expect(validatePackage(reopened.opcPackage)
      .filter(({ severity }) => severity === 'error')).toEqual([]);

    if (false) {
      const direction: TableCellTextDirection | undefined = table.textDirection;
      table.textDirection = 'horz';
      table.textDirection = 'vert';
      table.textDirection = 'vert270';
      table.textDirection = 'wordArtVert';
      table.textDirection = undefined;
      // @ts-expect-error unsupported table-level text direction
      table.textDirection = 'eaVert';
      void direction;
    }
  });

  it('rejects invalid table-cell vertical alignments and physical coordinates before mutation', async () => {
    const document = await PptxDocument.open(await tableTextDirectionFixture());
    const slide = document.slides[0]!;
    const table = slide.shapes[0] as TableModel;
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    const alignments = table.rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment));
    const fits = table.rows.map(({ cells }) => cells.map(({ textFit }) => textFit));
    const directions = table.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection));
    const text = table.rows.map(({ cells }) => cells.map((cell) => cell.text));

    const invalidValues = [
      null,
      false,
      true,
      0,
      1,
      '',
      'Middle',
      ' middle ',
      'mid',
      'center',
      'just',
      'dist',
      {},
      [],
      Symbol('vertical alignment'),
    ];
    for (const value of invalidValues) {
      expect(() => table.setCellVerticalAlignment(2, 0, value as never)).toThrow(TypeError);
    }
    const invalidCoordinates = [
      [-1, 0],
      [0, -1],
      [0.5, 0],
      [0, 0.5],
      [Number.NaN, 0],
      [0, Number.NaN],
      [Number.POSITIVE_INFINITY, 0],
      [0, Number.NEGATIVE_INFINITY],
      [3, 0],
      [2, 5],
    ];
    for (const [row, column] of invalidCoordinates) {
      expect(() => table.setCellVerticalAlignment(row!, column!, 'middle')).toThrow(RangeError);
    }

    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(document.slides).toHaveLength(1);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment))).toEqual(alignments);
    expect(table.rows.map(({ cells }) => cells.map(({ textFit }) => textFit))).toEqual(fits);
    expect(table.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection))).toEqual(directions);
    expect(table.rows.map(({ cells }) => cells.map((cell) => cell.text))).toEqual(text);
  });

  it('edits table-cell margins through duplicate, rollback, and reopen lifecycles', async () => {
    const document = await PptxDocument.open(await tableTextDirectionFixture());
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
    const slide = document.slides[0]!;
    const table = slide.shapes[0] as TableModel;
    const originalMargins = [
      { top: 4, right: 8, bottom: 12, left: 16 },
      { top: 0, left: 2 },
      { right: 7.2 },
      undefined,
      { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 },
    ];
    expect(table.rows[2]!.cells.map(({ margins }) => margins)).toEqual(originalMargins);
    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;

    table.setCellMargins(2, 0, 6);
    table.setCellMargins(2, 1, [1, 2, 3, 4]);
    table.setCellMargins(2, 2, undefined);
    table.setCellMargins(2, 2, { top: 5, left: 7 });
    table.setCellMargins(2, 3, 0);
    table.setCellMargins(2, 4, { bottom: 9 });
    table.setCellTextFit(2, 0, 'resize');
    table.setCellTextDirection(2, 3, 'horz');
    table.setCellVerticalAlignment(2, 1, 'bottom');
    table.setCellText(2, 4, 'Edited merged margins');
    table.setTransform({ x: inches(2) });
    const snapshot = table.rows;
    (snapshot[2]!.cells[1]!.margins as { top?: number }).top = 99;

    const editedMargins = [
      { top: 6, right: 6, bottom: 6, left: 6 },
      { top: 1, right: 2, bottom: 3, left: 4 },
      { top: 5, left: 7 },
      { top: 0, right: 0, bottom: 0, left: 0 },
      { bottom: 9 },
    ];
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[2]!.cells.map(({ margins }) => margins)).toEqual(editedMargins);
    expect(table.rows[2]!.cells.map(({ textFit }) => textFit)).toEqual([
      'resize',
      'shrink',
      'resize',
      undefined,
      undefined,
    ]);
    expect(table.rows[2]!.cells.map(({ textDirection }) => textDirection)).toEqual([
      'horz',
      'vert',
      'vert270',
      'horz',
      'wordArtVert',
    ]);
    expect(table.rows[2]!.cells.map(({ verticalAlignment }) => verticalAlignment)).toEqual([
      'top',
      'bottom',
      'bottom',
      undefined,
      'middle',
    ]);
    expect(table.rows[2]!.cells[4]!.text).toBe('Edited merged margins');
    const editedXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(editedXml).toContain('<a:tcPr vert="horz" anchor="t" marT="76200" marR="76200" marB="76200" marL="76200" keep="FIT-NONE"><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></a:tcPr>');
    expect(editedXml).toContain('<a:tcPr vert="vert" anchor="b" marT="12700" marL="50800" keep="FIT-SHRINK" marR="25400" marB="38100"><a:solidFill><a:schemeClr val="accent2"><a:alpha val="75000"/></a:schemeClr></a:solidFill></a:tcPr>');
    expect(editedXml).toContain('<a:tcPr vert="vert270" anchor="b" keep="FIT-RESIZE" marL="88900" marT="63500"><a:noFill/></a:tcPr>');
    expect(editedXml).toContain('<a:tcPr keep="FIT-ABSENT" marL="0" marR="0" marT="0" marB="0" vert="horz"/>');
    expect(editedXml).toContain('<a:tcPr vert="wordArtVert" anchor="ctr" marB="114300" keep="FIT-MERGED"><a:solidFill><a:srgbClr val="70AD47"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:tcPr>');
    expect(editedXml).toContain('<a:bodyPr custom="NONE"><a:spAutoFit/></a:bodyPr>');
    expect(editedXml).toContain('<a:off x="1828800" y="914400"/>');
    expect(duplicateTable.rows[2]!.cells.map(({ margins }) => margins)).toEqual(originalMargins);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    const rollbackJournal = [...document.opcPackage.mutations];
    expect(() =>
      document.transaction(() => {
        table.setCellMargins(2, 0, 1);
        table.setCellMargins(2, 3, undefined);
        throw new Error('restore public table margin edits');
      }),
    ).toThrow('restore public table margin edits');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[2]!.cells.map(({ margins }) => margins)).toEqual(editedMargins);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedEdited = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(reopenedEdited.rows[2]!.cells.map(({ margins }) => margins)).toEqual(editedMargins);
    expect(reopenedDuplicate.rows[2]!.cells.map(({ margins }) => margins)).toEqual(originalMargins);
  });

  it('rejects invalid table-cell margins and physical coordinates before mutation', async () => {
    const document = await PptxDocument.open(await tableTextDirectionFixture());
    const slide = document.slides[0]!;
    const table = slide.shapes[0] as TableModel;
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    const margins = table.rows.map(({ cells }) => cells.map((cell) => cell.margins));
    const alignments = table.rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment));
    const fits = table.rows.map(({ cells }) => cells.map(({ textFit }) => textFit));
    const directions = table.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection));
    const text = table.rows.map(({ cells }) => cells.map((cell) => cell.text));

    const invalidValues = [
      null,
      false,
      true,
      '',
      '4',
      [],
      [1, 2, 3],
      [1, 2, 3, 4, 5],
      [1, 2, Number.NaN, 4],
      { top: '4' },
      { top: 1, extra: 2 },
      Number.NaN,
      Number.POSITIVE_INFINITY,
      Number.NEGATIVE_INFINITY,
      (2_147_483_647 + 1) / 12_700,
      (-2_147_483_648 - 1) / 12_700,
      Symbol('table cell margins'),
    ];
    for (const value of invalidValues) {
      expect(() => table.setCellMargins(2, 0, value as never)).toThrow();
    }
    const invalidCoordinates = [
      [-1, 0],
      [0, -1],
      [0.5, 0],
      [0, 0.5],
      [Number.NaN, 0],
      [0, Number.NaN],
      [Number.POSITIVE_INFINITY, 0],
      [0, Number.NEGATIVE_INFINITY],
      [3, 0],
      [2, 5],
    ];
    for (const [row, column] of invalidCoordinates) {
      expect(() => table.setCellMargins(row!, column!, 4)).toThrow(RangeError);
    }

    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(document.slides).toHaveLength(1);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows.map(({ cells }) => cells.map((cell) => cell.margins))).toEqual(margins);
    expect(table.rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment))).toEqual(alignments);
    expect(table.rows.map(({ cells }) => cells.map(({ textFit }) => textFit))).toEqual(fits);
    expect(table.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection))).toEqual(directions);
    expect(table.rows.map(({ cells }) => cells.map((cell) => cell.text))).toEqual(text);
  });

  it('edits table-cell fills through duplicate, rollback, and reopen lifecycles', async () => {
    const document = await PptxDocument.open(await tableTextDirectionFixture());
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
    const slide = document.slides[0]!;
    const table = slide.shapes[0] as TableModel;
    const originalFills = [
      { kind: 'solid', color: { kind: 'srgb', value: '4472C4' } },
      {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 25,
      },
      { kind: 'none' },
      undefined,
      {
        kind: 'solid',
        color: { kind: 'srgb', value: '70AD47' },
        transparency: 50,
      },
    ];
    expect(table.rows[2]!.cells.map(({ fill }) => fill)).toEqual(originalFills);
    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;

    table.setCellFill(2, 0, {
      kind: 'solid',
      color: { kind: 'srgb', value: '#C00000' },
    });
    table.setCellFill(2, 1, {
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 0,
    });
    table.setCellFill(2, 2, {
      kind: 'solid',
      color: { kind: 'srgb', value: '0000FF' },
      transparency: 33.333,
    });
    table.setCellFill(2, 2, { kind: 'none' });
    table.setCellFill(2, 3, {
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent4' },
    });
    table.setCellFill(2, 3, undefined);
    table.setCellFill(2, 4, {
      kind: 'solid',
      color: { kind: 'srgb', value: '00B0F0' },
      transparency: 100,
    });
    table.setCellMargins(2, 0, 6);
    table.setCellTextFit(2, 1, 'resize');
    table.setCellTextDirection(2, 3, 'horz');
    table.setCellVerticalAlignment(2, 1, 'bottom');
    table.setCellText(2, 4, 'Edited merged fill');
    table.setTransform({ x: inches(2) });
    const snapshot = table.rows;
    const mutableFill = snapshot[2]!.cells[1]!.fill as {
      kind: string;
      color: { kind: string; value: string };
      transparency?: number;
    };
    mutableFill.kind = 'none';
    mutableFill.color.value = 'FFFFFF';
    mutableFill.transparency = 99;

    const editedFills = [
      { kind: 'solid', color: { kind: 'srgb', value: 'C00000' } },
      {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 0,
      },
      { kind: 'none' },
      undefined,
      {
        kind: 'solid',
        color: { kind: 'srgb', value: '00B0F0' },
        transparency: 100,
      },
    ];
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[2]!.cells.map(({ fill }) => fill)).toEqual(editedFills);
    expect(table.rows[2]!.cells[0]!.margins).toEqual({
      top: 6,
      right: 6,
      bottom: 6,
      left: 6,
    });
    expect(table.rows[2]!.cells[1]!.textFit).toBe('resize');
    expect(table.rows[2]!.cells[1]!.verticalAlignment).toBe('bottom');
    expect(table.rows[2]!.cells[3]!.textDirection).toBe('horz');
    expect(table.rows[2]!.cells[4]!.text).toBe('Edited merged fill');
    const editedXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(editedXml).toContain(
      '<a:tcPr vert="horz" anchor="t" marT="76200" marR="76200" marB="76200" marL="76200" keep="FIT-NONE"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:tcPr>',
    );
    expect(editedXml).toContain(
      '<a:tcPr vert="vert" anchor="b" marT="0" marL="25400" keep="FIT-SHRINK"><a:solidFill><a:schemeClr val="accent1"><a:alpha val="100000"/></a:schemeClr></a:solidFill></a:tcPr>',
    );
    expect(editedXml).toContain(
      '<a:tcPr vert="vert270" anchor="b" marR="91440" keep="FIT-RESIZE"><a:noFill/></a:tcPr>',
    );
    expect(editedXml).toContain('<a:tcPr keep="FIT-ABSENT" vert="horz"></a:tcPr>');
    expect(editedXml).toContain(
      '<a:tcPr vert="wordArtVert" anchor="ctr" marT="45720" marR="91440" marB="45720" marL="91440" keep="FIT-MERGED"><a:solidFill><a:srgbClr val="00B0F0"><a:alpha val="0"/></a:srgbClr></a:solidFill></a:tcPr>',
    );
    expect(editedXml).toContain('<a:off x="1828800" y="914400"/>');
    expect(duplicateTable.rows[2]!.cells.map(({ fill }) => fill)).toEqual(originalFills);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    const rollbackJournal = [...document.opcPackage.mutations];
    expect(() =>
      document.transaction(() => {
        table.setCellFill(2, 0, { kind: 'none' });
        table.setCellFill(2, 3, {
          kind: 'solid',
          color: { kind: 'srgb', value: 'FFFFFF' },
        });
        throw new Error('restore public table fill edits');
      }),
    ).toThrow('restore public table fill edits');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[2]!.cells.map(({ fill }) => fill)).toEqual(editedFills);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedEdited = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(reopenedEdited.rows[2]!.cells.map(({ fill }) => fill)).toEqual(editedFills);
    expect(reopenedDuplicate.rows[2]!.cells.map(({ fill }) => fill)).toEqual(originalFills);
  });

  it('rejects invalid table-cell fills and physical coordinates before mutation', async () => {
    const document = await PptxDocument.open(await tableTextDirectionFixture());
    const slide = document.slides[0]!;
    const table = slide.shapes[0] as TableModel;
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    const fills = table.rows.map(({ cells }) => cells.map((cell) => cell.fill));
    const margins = table.rows.map(({ cells }) => cells.map((cell) => cell.margins));
    const alignments = table.rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment));
    const fits = table.rows.map(({ cells }) => cells.map(({ textFit }) => textFit));
    const directions = table.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection));
    const text = table.rows.map(({ cells }) => cells.map((cell) => cell.text));

    const invalidValues = [
      null,
      false,
      true,
      '',
      [],
      {},
      { kind: 'none', color: { kind: 'srgb', value: 'FF0000' } },
      { kind: 'unknown' },
      { kind: 'solid' },
      { kind: 'solid', color: null },
      { kind: 'solid', color: { kind: 'srgb', value: 'FFF' } },
      { kind: 'solid', color: { kind: 'srgb', value: 'GG0000' } },
      { kind: 'solid', color: { kind: 'scheme', value: 'unknown' } },
      { kind: 'solid', color: { kind: 'srgb', value: 'FF0000', extra: true } },
      { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: -0.001 },
      { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: 100.001 },
      { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: Number.NaN },
      { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: Number.POSITIVE_INFINITY },
      Symbol('table cell fill'),
    ];
    for (const value of invalidValues) {
      expect(() => table.setCellFill(2, 0, value as never)).toThrow();
    }
    const invalidCoordinates = [
      [-1, 0],
      [0, -1],
      [0.5, 0],
      [0, 0.5],
      [Number.NaN, 0],
      [0, Number.NaN],
      [Number.POSITIVE_INFINITY, 0],
      [0, Number.NEGATIVE_INFINITY],
      [3, 0],
      [2, 5],
    ];
    for (const [row, column] of invalidCoordinates) {
      expect(() => table.setCellFill(row!, column!, { kind: 'none' })).toThrow(RangeError);
    }

    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(document.slides).toHaveLength(1);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows.map(({ cells }) => cells.map((cell) => cell.fill))).toEqual(fills);
    expect(table.rows.map(({ cells }) => cells.map((cell) => cell.margins))).toEqual(margins);
    expect(table.rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment))).toEqual(alignments);
    expect(table.rows.map(({ cells }) => cells.map(({ textFit }) => textFit))).toEqual(fits);
    expect(table.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection))).toEqual(directions);
    expect(table.rows.map(({ cells }) => cells.map((cell) => cell.text))).toEqual(text);
  });

  it('edits table-cell borders through duplicate, rollback, and reopen lifecycles', async () => {
    const document = await PptxDocument.open(await tableBordersFixture());
    expect(validatePackage(document.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
    const slide = document.slides[0]!;
    const table = slide.shapes[0] as TableModel;
    const line = (
      color: { kind: 'srgb' | 'scheme'; value: string },
      width: number,
      style?: 'solid' | 'dash',
    ) => ({ kind: 'line' as const, color, width, ...(style ? { style } : {}) });
    const red = line({ kind: 'srgb', value: 'C00000' }, 1, 'solid');
    const originalBorders = [
      { top: red, right: red, bottom: red, left: red },
      {
        top: line({ kind: 'scheme', value: 'accent1' }, 1.5, 'dash'),
        right: line({ kind: 'srgb', value: '00FF00' }, 0),
        bottom: { kind: 'none' },
      },
      {
        top: { kind: 'none' },
        right: { kind: 'none' },
        bottom: { kind: 'none' },
        left: { kind: 'none' },
      },
      undefined,
      { left: line({ kind: 'srgb', value: '333333' }, 2, 'solid') },
    ];
    expect(table.rows[2]!.cells.map(({ borders }) => borders)).toEqual(originalBorders);
    const duplicate = document.duplicateSlide(0);
    const duplicateTable = duplicate.shapes[0] as TableModel;

    table.setCellBorders(2, 0, {
      kind: 'line',
      color: { kind: 'scheme', value: 'accent3' },
      width: 2.25,
      style: 'dash',
    });
    table.setCellBorders(2, 1, [
      {
        kind: 'line',
        color: { kind: 'srgb', value: '#112233' },
        width: 2,
      },
      { kind: 'none' },
      {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent4' },
        width: 0,
        style: 'solid',
      },
      undefined,
    ]);
    table.setCellBorders(2, 2, {
      top: {
        kind: 'line',
        color: { kind: 'srgb', value: '0000FF' },
        width: 3,
        style: 'solid',
      },
      left: { kind: 'none' },
    });
    table.setCellBorders(2, 3, { kind: 'none' });
    table.setCellBorders(2, 3, undefined);
    table.setCellBorders(2, 4, {
      kind: 'line',
      color: { kind: 'srgb', value: '00B0F0' },
      width: 0.5,
    });
    table.setCellMargins(2, 0, 6);
    table.setCellFill(2, 1, { kind: 'solid', color: { kind: 'scheme', value: 'accent5' } });
    table.setCellTextFit(2, 1, 'resize');
    table.setCellTextDirection(2, 3, 'horz');
    table.setCellVerticalAlignment(2, 1, 'bottom');
    table.setCellText(2, 4, 'Edited merged borders');
    table.setTransform({ x: inches(2) });

    const snapshot = table.rows;
    const mutable = snapshot[2]!.cells[0]!.borders as {
      top?: {
        kind: string;
        color: { kind: string; value: string };
        width: number;
        style?: string;
      };
    };
    mutable.top!.kind = 'none';
    mutable.top!.color.value = 'FFFFFF';
    mutable.top!.width = 99;
    mutable.top!.style = 'solid';

    const accent3 = line({ kind: 'scheme', value: 'accent3' }, 2.25, 'dash');
    const cyan = line({ kind: 'srgb', value: '00B0F0' }, 0.5);
    const editedBorders = [
      { top: accent3, right: accent3, bottom: accent3, left: accent3 },
      {
        top: line({ kind: 'srgb', value: '112233' }, 2),
        right: { kind: 'none' },
        bottom: line({ kind: 'scheme', value: 'accent4' }, 0, 'solid'),
      },
      {
        top: line({ kind: 'srgb', value: '0000FF' }, 3, 'solid'),
        left: { kind: 'none' },
      },
      undefined,
      { top: cyan, right: cyan, bottom: cyan, left: cyan },
    ];
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[2]!.cells.map(({ borders }) => borders)).toEqual(editedBorders);
    expect(table.rows[2]!.cells[0]!.margins).toEqual({
      top: 6,
      right: 6,
      bottom: 6,
      left: 6,
    });
    expect(table.rows[2]!.cells[1]!.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent5' },
    });
    expect(table.rows[2]!.cells[1]!.textFit).toBe('resize');
    expect(table.rows[2]!.cells[1]!.verticalAlignment).toBe('bottom');
    expect(table.rows[2]!.cells[3]!.textDirection).toBe('horz');
    expect(table.rows[2]!.cells[4]!.text).toBe('Edited merged borders');
    const editedXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(editedXml).toContain(
      '<a:lnL w="28575" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="accent3"/></a:solidFill><a:prstDash val="sysDash"/>',
    );
    expect(editedXml).toContain(
      '<a:lnL w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnL><a:lnT w="38100"',
    );
    expect(editedXml).toContain('<a:lnTlToBr w="12700">');
    expect(editedXml).toContain('<a:off x="1828800" y="914400"/>');
    expect(duplicateTable.rows[2]!.cells.map(({ borders }) => borders)).toEqual(originalBorders);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    const rollbackJournal = [...document.opcPackage.mutations];
    expect(() =>
      document.transaction(() => {
        table.setCellBorders(2, 0, { kind: 'none' });
        table.setCellBorders(2, 3, {
          kind: 'line',
          color: { kind: 'srgb', value: 'FFFFFF' },
          width: 1,
        });
        throw new Error('restore public table border edits');
      }),
    ).toThrow('restore public table border edits');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(rollbackJournal);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[2]!.cells.map(({ borders }) => borders)).toEqual(editedBorders);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedEdited = reopened.slides[0]!.shapes[0] as TableModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
    expect(reopenedEdited.rows[2]!.cells.map(({ borders }) => borders)).toEqual(editedBorders);
    expect(reopenedDuplicate.rows[2]!.cells.map(({ borders }) => borders)).toEqual(originalBorders);
  });

  it('rejects invalid table-cell borders and physical coordinates before mutation', async () => {
    const document = await PptxDocument.open(await tableBordersFixture());
    const slide = document.slides[0]!;
    const table = slide.shapes[0] as TableModel;
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    const borders = table.rows.map(({ cells }) => cells.map((cell) => cell.borders));
    const fills = table.rows.map(({ cells }) => cells.map((cell) => cell.fill));
    const margins = table.rows.map(({ cells }) => cells.map((cell) => cell.margins));
    const alignments = table.rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment));
    const fits = table.rows.map(({ cells }) => cells.map(({ textFit }) => textFit));
    const directions = table.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection));
    const text = table.rows.map(({ cells }) => cells.map((cell) => cell.text));

    const sparse = Array(4);
    sparse[0] = { kind: 'none' };
    const extraTuple = Object.assign(
      [{ kind: 'none' }, undefined, undefined, undefined],
      { extra: true },
    );
    const invalidValues = [
      null,
      false,
      true,
      '',
      [],
      [{ kind: 'none' }, undefined, undefined],
      sparse,
      extraTuple,
      { top: undefined, extra: true },
      { kind: 'none', width: 1 },
      { kind: 'unknown' },
      { kind: 'line' },
      { kind: 'line', color: null, width: 1 },
      { kind: 'line', color: { kind: 'srgb', value: 'FFF' }, width: 1 },
      { kind: 'line', color: { kind: 'srgb', value: 'GG0000' }, width: 1 },
      { kind: 'line', color: { kind: 'scheme', value: 'unknown' }, width: 1 },
      { kind: 'line', color: { kind: 'srgb', value: 'FF0000', extra: true }, width: 1 },
      { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: -0.001 },
      { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 1584.001 },
      { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: Number.NaN },
      { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: Number.POSITIVE_INFINITY },
      { kind: 'line', color: { kind: 'srgb', value: 'FF0000' }, width: 1, style: 'dot' },
      Symbol('table cell borders'),
    ];
    for (const value of invalidValues) {
      expect(() => table.setCellBorders(2, 0, value as never)).toThrow();
    }
    const invalidCoordinates = [
      [-1, 0],
      [0, -1],
      [0.5, 0],
      [0, 0.5],
      [Number.NaN, 0],
      [0, Number.NaN],
      [Number.POSITIVE_INFINITY, 0],
      [0, Number.NEGATIVE_INFINITY],
      [3, 0],
      [2, 5],
    ];
    for (const [row, column] of invalidCoordinates) {
      expect(() => table.setCellBorders(row!, column!, { kind: 'none' })).toThrow(RangeError);
    }

    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(document.slides).toHaveLength(1);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows.map(({ cells }) => cells.map((cell) => cell.borders))).toEqual(borders);
    expect(table.rows.map(({ cells }) => cells.map((cell) => cell.fill))).toEqual(fills);
    expect(table.rows.map(({ cells }) => cells.map((cell) => cell.margins))).toEqual(margins);
    expect(table.rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment))).toEqual(alignments);
    expect(table.rows.map(({ cells }) => cells.map(({ textFit }) => textFit))).toEqual(fits);
    expect(table.rows.map(({ cells }) => cells.map(({ textDirection }) => textDirection))).toEqual(directions);
    expect(table.rows.map(({ cells }) => cells.map((cell) => cell.text))).toEqual(text);
  });

  it('exports table colspan, rowspan, merge snapshots, and editors through the SDK', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const spanOptions: AddTableCellOptions = { colspan: 2, rowspan: 1 };
    const mixed = slide.addTable([
      [
        { text: 'Horizontal', options: spanOptions },
        { text: 'Vertical', options: { rowspan: 2 } },
        'Right',
      ],
      ['Lower left', 'Lower middle', 'Lower right'],
    ], { name: 'SDK mixed spans' });
    const ordinary = slide.addTable([
      ['A', 'B'],
      ['C', 'D'],
    ], { name: 'SDK editable merge' });

    expect(mixed.rows.map(({ cells }) => cells.map(({ text }) => text))).toEqual([
      ['Horizontal', '', 'Vertical', 'Right'],
      ['Lower left', 'Lower middle', '', 'Lower right'],
    ]);
    expect(mixed.mergeRegions).toEqual([
      { rowIndex: 0, columnIndex: 0, rowspan: 1, colspan: 2 },
      { rowIndex: 0, columnIndex: 2, rowspan: 2, colspan: 1 },
    ]);
    const region: Readonly<TableMergeRegion> = mixed.mergeRegions![1]!;
    const cellMerge: Readonly<TableCellMerge> = mixed.rows[1]!.cells[2]!.merge!;
    expect(region).toEqual({ rowIndex: 0, columnIndex: 2, rowspan: 2, colspan: 1 });
    expect(cellMerge).toEqual({
      rowIndex: 0,
      columnIndex: 2,
      rowspan: 2,
      colspan: 1,
      isAnchor: false,
    });

    ordinary.mergeCells(0, 0, 2, 2);
    expect(slide.shapes.find(({ id }) => id === ordinary.id)).toBe(ordinary);
    expect(ordinary.mergeRegions).toEqual([
      { rowIndex: 0, columnIndex: 0, rowspan: 2, colspan: 2 },
    ]);
    ordinary.unmergeCell(1, 1);
    expect(slide.shapes.find(({ id }) => id === ordinary.id)).toBe(ordinary);
    expect(ordinary.mergeRegions).toEqual([]);

    const beforeFraction = await sdkPackageSnapshot(document);
    expect(() => ordinary.mergeCells(0, 0, 1.5, 2)).toThrow(RangeError);
    expect(await sdkPackageSnapshot(document)).toEqual(beforeFraction);

    if (false) {
      const wrongCasing: AddTableCellOptions = {
        // @ts-expect-error public span fields use exact PptxGenJS lowercase spelling.
        colSpan: 2,
      };
      const wrongRowCasing: AddTableCellOptions = {
        // @ts-expect-error public span fields use exact PptxGenJS lowercase spelling.
        rowSpan: 2,
      };
      const stringSpan: AddTableCellOptions = {
        // @ts-expect-error spans are numbers.
        colspan: '2',
      };
      const logicalOnly: TableMergeRegion = {
        rowIndex: 0,
        columnIndex: 0,
        rowspan: 2,
        colspan: 2,
        // @ts-expect-error snapshots expose physical coordinates only.
        logicalRowIndex: 0,
      };
      // @ts-expect-error merge snapshots are readonly.
      region.rowspan = 3;
      // @ts-expect-error mergeCells requires four numeric arguments.
      ordinary.mergeCells(0, 0, 2);
      // @ts-expect-error unmergeCell requires both physical coordinates.
      ordinary.unmergeCell(0);
      // @ts-expect-error mergeCells coordinates are numeric.
      ordinary.mergeCells('0', 0, 1, 2);
      void [wrongCasing, wrongRowCasing, stringSpan, logicalOnly];
    }

    const reopened = await PptxDocument.open(await document.write());
    const reopenedMixed = reopened.slides[0]!.shapes.find((shape): shape is TableModel =>
      shape instanceof TableModel && shape.name === 'SDK mixed spans')!;
    const reopenedOrdinary = reopened.slides[0]!.shapes.find(
      (shape): shape is TableModel =>
        shape instanceof TableModel && shape.name === 'SDK editable merge',
    )!;
    expect(reopenedMixed.mergeRegions).toEqual(mixed.mergeRegions);
    expect(reopenedMixed.rows).toEqual(mixed.rows);
    expect(reopenedOrdinary.mergeRegions).toEqual([]);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) =>
      severity === 'error')).toEqual([]);
  });

  it('exports table row and column structure editing through all presentation formats', async () => {
    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const document = PptxDocument.create({ format });
      const slide = document.addSlide();
      const table = slide.addTable([
        [`${format} A0`, `${format} A1`, `${format} A2`],
        [`${format} B0`, `${format} B1`, `${format} B2`],
        [`${format} C0`, `${format} C1`, `${format} C2`],
      ], {
        name: `SDK table structure ${format}`,
        columnWidths: [100, 200, 300],
        rowHeights: [10, 20, 30],
      });
      table.mergeCells(0, 0, 2, 2);
      const rowOptions: InsertTableRowsOptions = { rowHeights: 11 };
      const columnOptions: InsertTableColumnsOptions = { columnWidths: 21 };
      table.insertRows(1, rowOptions);
      table.insertColumns(1, columnOptions);
      table.setCellText(1, 1, `Hidden ${format}`);
      table.deleteRows(3);
      table.deleteColumns(3);
      expect(table.rowHeights).toEqual([10, 11, 20]);
      expect(table.columnWidths).toEqual([100, 21, 200]);
      expect(table.mergeRegions).toEqual([
        { rowIndex: 0, columnIndex: 0, rowspan: 3, colspan: 3 },
      ]);

      const duplicate = document.duplicateSlide(0);
      expect((duplicate.shapes[0] as TableModel).rows[1]!.cells[1]!.text)
        .toBe(`Hidden ${format}`);
      const reopened = await PptxDocument.open(await document.write());
      expect(reopened.format).toBe(format);
      const reopenedSource = reopened.slides[0]!.shapes[0] as TableModel;
      const reopenedDuplicate = reopened.slides[1]!.shapes[0] as TableModel;
      for (const candidate of [reopenedSource, reopenedDuplicate]) {
        expect(candidate.rowHeights).toEqual([10, 11, 20]);
        expect(candidate.columnWidths).toEqual([100, 21, 200]);
        expect(candidate.mergeRegions).toEqual([
          { rowIndex: 0, columnIndex: 0, rowspan: 3, colspan: 3 },
        ]);
        expect(candidate.rows[1]!.cells[1]!.text).toBe(`Hidden ${format}`);
      }
      expect(validatePackage(reopened.opcPackage).filter(
        ({ severity }) => severity === 'error',
      )).toEqual([]);
    }

    if (false) {
      const document = PptxDocument.create();
      const table = document.addSlide().addTable([['A', 'B'], ['C', 'D']]);
      const wrongRowName: InsertTableRowsOptions = {
        // @ts-expect-error row insert uses rowHeights.
        rowHeight: 1,
      };
      const wrongColumnName: InsertTableColumnsOptions = {
        // @ts-expect-error column insert uses columnWidths.
        columnWidth: 1,
      };
      // @ts-expect-error insertRows requires a physical row index.
      table.insertRows();
      // @ts-expect-error insertColumns index is numeric.
      table.insertColumns('0');
      // @ts-expect-error deleteRows count is numeric.
      table.deleteRows(0, '1');
      // @ts-expect-error deleteColumns accepts at most index and count.
      table.deleteColumns(0, 1, 2);
      void [wrongRowName, wrongColumnName];
    }
  });

  it('creates, edits, and round-trips plain-text paragraphs with normalized line endings', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const text = slide.addText(' First & <\r\nSecond\r\rFourth > \n');
    const normalized = ' First & <\nSecond\n\nFourth > \n';

    expect(text.text).toBe(normalized);
    let slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml.match(/<a:p>/g)).toHaveLength(5);
    expect(slideXml).toContain('<a:t xml:space="preserve"> First &amp; &lt;</a:t>');
    expect(slideXml).toContain('<a:t xml:space="preserve">Fourth &gt; </a:t>');

    text.text = 'Updated\r\n\rEnd';
    expect(text.text).toBe('Updated\n\nEnd');
    expect(slide.shapes[0]).toBe(text);
    slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml.match(/<a:p>/g)).toHaveLength(3);

    expect(() =>
      document.transaction(() => {
        text.text = 'Rollback\ntext';
        throw new Error('restore paragraphs');
      }),
    ).toThrow('restore paragraphs');
    expect(text.text).toBe('Updated\n\nEnd');

    const reopened = await PptxDocument.open(await document.write());
    expect((reopened.slides[0]!.shapes[0] as ShapeModel).text).toBe('Updated\n\nEnd');
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);

    const emptyDocument = PptxDocument.create();
    const empty = emptyDocument.addSlide().addText('');
    expect(empty.text).toBe('');
    empty.text = 'Filled from an empty paragraph';
    expect(empty.text).toBe('Filled from an empty paragraph');
  });

  it('creates, replaces, rolls back, and round-trips paragraph alignment', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const plain = slide.addText('Centered\n\nParagraphs', { align: 'center' });
    const input = [
      { runs: [{ text: 'Left' }], align: 'left' as const },
      { runs: [{ text: 'Default right' }] },
      { runs: [], align: 'justify' as const },
    ];
    const rich = slide.addRichText(input, { align: 'right' });

    expect(plain.richText.map(({ align }) => align)).toEqual(['center', 'center', 'center']);
    expect(rich.richText.map(({ align }) => align)).toEqual(['left', 'right', 'justify']);
    let slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml.match(/algn="ctr"/g)).toHaveLength(3);
    expect(slideXml).toContain('<a:pPr algn="l" indent="0" marL="0">');
    expect(slideXml).toContain('<a:pPr algn="r" indent="0" marL="0">');
    expect(slideXml).toContain('<a:pPr algn="just" indent="0" marL="0">');

    rich.richText = [
      { runs: rich.richText[0]!.runs, align: 'center' },
      { runs: rich.richText[1]!.runs },
      { runs: rich.richText[2]!.runs, align: 'right' },
    ];
    expect(rich.richText.map(({ align }) => align)).toEqual(['center', undefined, 'right']);
    slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml).not.toContain('<a:pPr algn="just"');

    expect(() =>
      document.transaction(() => {
        rich.richText = [{ runs: [{ text: 'Rollback' }], align: 'left' }];
        throw new Error('restore paragraph alignment');
      }),
    ).toThrow('restore paragraph alignment');
    expect(rich.richText.map(({ align }) => align)).toEqual(['center', undefined, 'right']);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedPlain = reopened.slides[0]!.shapes[0] as ShapeModel;
    const reopenedRich = reopened.slides[0]!.shapes[1] as ShapeModel;
    expect(reopenedPlain.richText.map(({ align }) => align)).toEqual(['center', 'center', 'center']);
    expect(reopenedRich.richText.map(({ align }) => align)).toEqual(['center', undefined, 'right']);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('publishes TEXT_ALIGNMENTS through the SDK lifecycle', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const created = TEXT_ALIGNMENTS.map((alignment) => {
      const typed: TextAlignment = alignment;
      return slide.addText(typed, { align: typed });
    });

    expect(created.map(({ richText }) => richText[0]?.align)).toEqual(TEXT_ALIGNMENTS);
    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides[0]?.shapes.map((shape) =>
      (shape as ShapeModel).richText[0]?.align)).toEqual(TEXT_ALIGNMENTS);
  });

  it('publishes TEXT_VERTICAL_ALIGNMENTS through text and table lifecycles', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const created = TEXT_VERTICAL_ALIGNMENTS.map((alignment) => {
      const typed: TextBoxVerticalAlignment = alignment;
      return slide.addText(typed, { valign: typed });
    });
    const table = slide.addTable([
      TEXT_VERTICAL_ALIGNMENTS.map((alignment) => ({
        text: alignment,
        options: { valign: alignment },
      })),
    ]);

    expect(created.map(({ verticalAlignment }) => verticalAlignment))
      .toEqual(TEXT_VERTICAL_ALIGNMENTS);
    expect(table.rows[0]?.cells.map(({ verticalAlignment }) => verticalAlignment))
      .toEqual(TEXT_VERTICAL_ALIGNMENTS);
    const journal = [...document.opcPackage.mutations];
    expect([...TEXT_VERTICAL_ALIGNMENTS]).toEqual(['top', 'middle', 'bottom']);
    expect(document.opcPackage.mutations).toEqual(journal);

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides[0]?.shapes.slice(0, 3).map((shape) =>
      (shape as ShapeModel).verticalAlignment)).toEqual(TEXT_VERTICAL_ALIGNMENTS);
    expect((reopened.slides[0]?.shapes[3] as TableModel).rows[0]?.cells.map(
      ({ verticalAlignment }) => verticalAlignment,
    )).toEqual(TEXT_VERTICAL_ALIGNMENTS);
  });

  it('creates, edits, duplicates, and reopens plain and rich paragraph RTL', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const omitted = slide.addText('Default direction', { name: 'RTL omitted' });
    const plain = slide.addText('مرحبا\n\nالعالم', {
      name: 'RTL plain',
      align: 'right',
      bullet: true,
      lang: 'ar-SA',
      level: 1,
      rtlMode: true,
      spacing: { before: 4, after: 6 },
      tabStops: [{ position: 1, alignment: 'right' }],
    });
    const explicitFalse = slide.addText('English', { name: 'RTL false', rtlMode: false });
    const rich = slide.addRichText([
      { runs: [{ text: 'Inherited' }] },
      { rtl: false, runs: [{ text: 'English override' }] },
      { rtl: true, runs: [] },
    ], {
      name: 'RTL rich',
      lang: 'he-IL',
      rtlMode: true,
    });

    expect(omitted.richText.map(({ rtl }) => rtl)).toEqual([undefined]);
    expect(plain.richText.map(({ rtl }) => rtl)).toEqual([true, true, true]);
    expect(explicitFalse.richText.map(({ rtl }) => rtl)).toEqual([false]);
    expect(rich.richText.map(({ rtl }) => rtl)).toEqual([true, false, true]);
    let slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml.match(/rtl="1"/g)).toHaveLength(5);
    expect(slideXml.match(/rtl="0"/g)).toHaveLength(2);
    expect(slideXml).toMatch(/<a:pPr algn="r"[^>]*rtl="1"[^>]*lvl="1"/);
    expect(slideXml).toContain('<a:bodyPr wrap="square" rtlCol="0"');
    const omittedStart = slideXml.indexOf('name="RTL omitted"');
    const omittedEnd = slideXml.indexOf('</p:sp>', omittedStart);
    expect(slideXml.slice(omittedStart, omittedEnd)).not.toContain(' rtl=');

    document.duplicateSlide(0);
    rich.richText = [
      { rtl: true, runs: [{ text: 'RTL' }] },
      { rtl: false, runs: [{ text: 'LTR' }] },
      { runs: [{ text: 'Cleared' }] },
    ];
    expect(rich.richText.map(({ rtl }) => rtl)).toEqual([true, false, undefined]);
    slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    const clearedText = slideXml.indexOf('>Cleared<', slideXml.indexOf('name="RTL rich"'));
    const clearedParagraphStart = slideXml.lastIndexOf('<a:p>', clearedText);
    const clearedParagraphEnd = slideXml.indexOf('</a:p>', clearedText);
    expect(slideXml.slice(clearedParagraphStart, clearedParagraphEnd)).not.toContain(' rtl=');

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    expect(() =>
      document.transaction(() => {
        rich.richText = [{ rtl: false, runs: [{ text: 'Rollback' }] }];
        throw new Error('restore paragraph RTL');
      }),
    ).toThrow('restore paragraph RTL');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(slide.shapes[3]).toBe(rich);

    const reopened = await PptxDocument.open(await document.write());
    const edited = reopened.slides[0]!.shapes[3] as ShapeModel;
    const duplicated = reopened.slides[1]!.shapes[3] as ShapeModel;
    expect(edited.richText.map(({ rtl }) => rtl)).toEqual([true, false, undefined]);
    expect(duplicated.richText.map(({ rtl }) => rtl)).toEqual([true, false, true]);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('rejects malformed paragraph RTL before changing slide package state', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addText('Original');
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    const invalidValues = [null, 0, 'true', {}, [], Symbol('rtl')] as const;

    for (const rtl of invalidValues) {
      expect(() => slide.addText('Invalid', { rtlMode: rtl as never })).toThrow(TypeError);
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid' }] }], {
        rtlMode: rtl as never,
      })).toThrow(TypeError);
      expect(() => slide.addRichText([{
        rtl: rtl as never,
        runs: [{ text: 'Invalid' }],
      }])).toThrow(TypeError);
      expect(() => {
        shape.richText = [{ rtl: rtl as never, runs: [{ text: 'Invalid' }] }];
      }).toThrow(TypeError);
    }

    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(slide.shapes[0]).toBe(shape);
  });

  it('creates, edits, duplicates, and reopens non-list paragraph left margins', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const omitted = slide.addText('Default margin');
    const plain = slide.addText('Twelve and a half\r\n\rSecond', {
      align: 'center',
      paragraphMarginLeft: 12.5,
      rtlMode: false,
      spacing: { after: 4 },
      tabStops: [{ position: 1.5 }],
    });
    const rich = slide.addRichText([
      { runs: [{ text: 'Default twenty-four' }] },
      { marginLeft: 12, runs: [{ text: 'Override twelve' }] },
      { bullet: true, marginLeft: false, runs: [{ text: 'Independent bullet' }] },
      { marginLeft: false, runs: [] },
    ], { paragraphMarginLeft: 24 });

    expect(omitted.richText.map(({ marginLeft }) => marginLeft)).toEqual([0]);
    expect(plain.richText.map(({ marginLeft }) => marginLeft)).toEqual([12.5, 12.5, 12.5]);
    expect(rich.richText.map(({ marginLeft }) => marginLeft)).toEqual([24, 12, undefined, undefined]);
    expect(rich.richText[2]!.bullet).toEqual({ kind: 'bullet', character: '•', indent: 27 });
    let slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml.match(/marL="158750"/g)).toHaveLength(3);
    expect(slideXml).toContain('marL="304800"');
    expect(slideXml).toContain('marL="152400"');
    expect(slideXml).toContain('indent="-342900" marL="342900"');

    document.duplicateSlide(0);
    plain.text = 'Updated\nCloned margin';
    expect(plain.richText.map(({ marginLeft }) => marginLeft)).toEqual([12.5, 12.5]);
    rich.richText = [
      { marginLeft: 6, runs: [{ text: 'Six' }] },
      { marginLeft: 0, runs: [{ text: 'Zero' }] },
      { marginLeft: false, runs: [{ text: 'Cleared false' }] },
      { runs: [{ text: 'Cleared omitted' }] },
    ];
    expect(rich.richText.map(({ marginLeft }) => marginLeft)).toEqual([6, 0, undefined, undefined]);
    slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml).toContain('marL="76200"');
    expect(slideXml).toContain('marL="0"');
    const clearedStart = slideXml.indexOf('>Cleared false<');
    const clearedEnd = slideXml.indexOf('</a:p>', clearedStart);
    expect(slideXml.slice(slideXml.lastIndexOf('<a:p>', clearedStart), clearedEnd)).not.toContain('marL=');

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    expect(() =>
      document.transaction(() => {
        rich.richText = [{ marginLeft: 48, runs: [{ text: 'Rollback' }] }];
        throw new Error('restore paragraph left margin');
      }),
    ).toThrow('restore paragraph left margin');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(slide.shapes[2]).toBe(rich);

    const reopened = await PptxDocument.open(await document.write());
    const editedPlain = reopened.slides[0]!.shapes[1] as ShapeModel;
    const editedRich = reopened.slides[0]!.shapes[2] as ShapeModel;
    const duplicatedRich = reopened.slides[1]!.shapes[2] as ShapeModel;
    expect(editedPlain.richText.map(({ marginLeft }) => marginLeft)).toEqual([12.5, 12.5]);
    expect(editedRich.richText.map(({ marginLeft }) => marginLeft)).toEqual([6, 0, undefined, undefined]);
    expect(duplicatedRich.richText.map(({ marginLeft }) => marginLeft)).toEqual([
      24,
      12,
      undefined,
      undefined,
    ]);
    expect(duplicatedRich.richText[2]!.bullet).toEqual({ kind: 'bullet', character: '•', indent: 27 });
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('rejects malformed or conflicting paragraph left margins before package mutation', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addText('Original');
    slide.addRichText([{ marginLeft: false, runs: [{ text: 'Suppressed margin' }] }], {
      bullet: true,
      paragraphMarginLeft: 12,
    });
    slide.addRichText([{ bullet: false, runs: [{ text: 'Suppressed bullet' }] }], {
      bullet: true,
      paragraphMarginLeft: 12,
    });
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    const invalidOuterTypes = [null, true, false, '12', {}, [], Symbol('margin'), Number.NaN, Infinity];
    const invalidParagraphTypes = [null, true, '12', {}, [], Symbol('margin'), Number.NaN, Infinity];
    const invalidRanges = [-0.01, 4032.01];

    for (const paragraphMarginLeft of invalidOuterTypes) {
      expect(() => slide.addText('Invalid', {
        paragraphMarginLeft: paragraphMarginLeft as never,
      })).toThrow(TypeError);
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid' }] }], {
        paragraphMarginLeft: paragraphMarginLeft as never,
      })).toThrow(TypeError);
    }
    for (const marginLeft of invalidParagraphTypes) {
      expect(() => slide.addRichText([{
        marginLeft: marginLeft as never,
        runs: [{ text: 'Invalid' }],
      }])).toThrow(TypeError);
      expect(() => {
        shape.richText = [{ marginLeft: marginLeft as never, runs: [{ text: 'Invalid' }] }];
      }).toThrow(TypeError);
    }
    for (const marginLeft of invalidRanges) {
      expect(() => slide.addText('Invalid', { paragraphMarginLeft: marginLeft })).toThrow(RangeError);
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid' }] }], {
        paragraphMarginLeft: marginLeft,
      })).toThrow(RangeError);
      expect(() => slide.addRichText([{
        marginLeft,
        runs: [{ text: 'Invalid' }],
      }])).toThrow(RangeError);
      expect(() => {
        shape.richText = [{ marginLeft, runs: [{ text: 'Invalid' }] }];
      }).toThrow(RangeError);
    }
    expect(() => slide.addText('Conflict', {
      bullet: true,
      paragraphMarginLeft: 12,
    })).toThrow(TypeError);
    expect(() => slide.addRichText([{ runs: [{ text: 'Conflict' }] }], {
      bullet: true,
      paragraphMarginLeft: 12,
    })).toThrow(TypeError);
    expect(() => slide.addRichText([{
      bullet: true,
      marginLeft: 12,
      runs: [{ text: 'Conflict' }],
    }])).toThrow(TypeError);

    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(slide.shapes[0]).toBe(shape);
  });

  it('creates, edits, duplicates, and reopens paragraph right margins with list coexistence', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const omitted = slide.addText('Default right margin');
    const plain = slide.addText('Twelve and a half\r\n\rSecond', {
      align: 'center',
      paragraphMarginLeft: 6,
      paragraphMarginRight: 12.5,
      rtlMode: false,
      spacing: { after: 4 },
      tabStops: [{ position: 1.5 }],
    });
    const rich = slide.addRichText([
      { runs: [{ text: 'Default twenty-four' }] },
      { marginRight: 12, runs: [{ text: 'Override twelve' }] },
      { marginRight: false, runs: [] },
      { bullet: true, runs: [{ text: 'Bullet default' }] },
      {
        bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
        marginRight: 18,
        runs: [{ text: 'Number override' }],
      },
    ], { paragraphMarginRight: 24 });

    expect(omitted.richText.map(({ marginRight }) => marginRight)).toEqual([undefined]);
    expect(plain.richText.map(({ marginLeft, marginRight }) => ({ marginLeft, marginRight }))).toEqual([
      { marginLeft: 6, marginRight: 12.5 },
      { marginLeft: 6, marginRight: 12.5 },
      { marginLeft: 6, marginRight: 12.5 },
    ]);
    expect(rich.richText.map(({ marginRight }) => marginRight)).toEqual([24, 12, undefined, 24, 18]);
    expect(rich.richText[3]!.bullet).toEqual({ kind: 'bullet', character: '•', indent: 27 });
    expect(rich.richText[4]!.bullet).toEqual({
      kind: 'number',
      style: 'romanUcPeriod',
      startAt: 3,
      indent: 22,
    });
    let slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml.match(/marR="158750"/g)).toHaveLength(3);
    expect(slideXml).toContain('marR="304800"');
    expect(slideXml).toContain('marR="152400"');
    expect(slideXml).toContain('marR="228600"');
    expect(slideXml).toContain('indent="-342900" marL="342900"');
    expect(slideXml).toContain('indent="-279400" marL="279400"');

    document.duplicateSlide(0);
    plain.text = 'Updated\nCloned margins';
    expect(plain.richText.map(({ marginRight }) => marginRight)).toEqual([12.5, 12.5]);
    rich.richText = [
      { marginRight: 6, runs: [{ text: 'Six' }] },
      { marginRight: 0, runs: [{ text: 'Zero' }] },
      { marginRight: false, runs: [{ text: 'Cleared false' }] },
      { runs: [{ text: 'Cleared omitted' }] },
      { bullet: true, marginRight: 9, runs: [{ text: 'Bullet nine' }] },
    ];
    expect(rich.richText.map(({ marginRight }) => marginRight)).toEqual([6, 0, undefined, undefined, 9]);
    slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml).toContain('marR="76200"');
    expect(slideXml).toContain('marR="0"');
    expect(slideXml).toContain('marR="114300"');
    const clearedStart = slideXml.indexOf('>Cleared false<');
    const clearedEnd = slideXml.indexOf('</a:p>', clearedStart);
    expect(slideXml.slice(slideXml.lastIndexOf('<a:p>', clearedStart), clearedEnd)).not.toContain('marR=');

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    expect(() =>
      document.transaction(() => {
        rich.richText = [{ marginRight: 48, runs: [{ text: 'Rollback' }] }];
        throw new Error('restore paragraph right margin');
      }),
    ).toThrow('restore paragraph right margin');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(slide.shapes[2]).toBe(rich);

    const reopened = await PptxDocument.open(await document.write());
    const editedPlain = reopened.slides[0]!.shapes[1] as ShapeModel;
    const editedRich = reopened.slides[0]!.shapes[2] as ShapeModel;
    const duplicatedRich = reopened.slides[1]!.shapes[2] as ShapeModel;
    expect(editedPlain.richText.map(({ marginRight }) => marginRight)).toEqual([12.5, 12.5]);
    expect(editedRich.richText.map(({ marginRight }) => marginRight)).toEqual([6, 0, undefined, undefined, 9]);
    expect(duplicatedRich.richText.map(({ marginRight }) => marginRight)).toEqual([
      24,
      12,
      undefined,
      24,
      18,
    ]);
    expect(duplicatedRich.richText[3]!.bullet).toEqual({ kind: 'bullet', character: '•', indent: 27 });
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('rejects malformed paragraph right margins before package mutation', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addText('Original');
    const boundary = slide.addText('Bullet boundary', { bullet: true, paragraphMarginRight: 4032 });
    slide.addRichText([{
      bullet: { kind: 'number' },
      marginRight: 18,
      runs: [{ text: 'Numbered margin' }],
    }], { paragraphMarginRight: 24 });
    expect(boundary.richText[0]!.marginRight).toBe(4032);
    expect(new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes))
      .toContain('marR="51206400"');
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    const invalidOuterTypes = [null, true, false, '12', {}, [], Symbol('margin'), Number.NaN, Infinity];
    const invalidParagraphTypes = [null, true, '12', {}, [], Symbol('margin'), Number.NaN, Infinity];
    const invalidRanges = [-0.01, 4032.01];

    for (const paragraphMarginRight of invalidOuterTypes) {
      expect(() => slide.addText('Invalid', {
        paragraphMarginRight: paragraphMarginRight as never,
      })).toThrow(TypeError);
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid' }] }], {
        paragraphMarginRight: paragraphMarginRight as never,
      })).toThrow(TypeError);
    }
    for (const marginRight of invalidParagraphTypes) {
      expect(() => slide.addRichText([{
        marginRight: marginRight as never,
        runs: [{ text: 'Invalid' }],
      }])).toThrow(TypeError);
      expect(() => {
        shape.richText = [{ marginRight: marginRight as never, runs: [{ text: 'Invalid' }] }];
      }).toThrow(TypeError);
    }
    for (const marginRight of invalidRanges) {
      expect(() => slide.addText('Invalid', { paragraphMarginRight: marginRight })).toThrow(RangeError);
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid' }] }], {
        paragraphMarginRight: marginRight,
      })).toThrow(RangeError);
      expect(() => slide.addRichText([{
        marginRight,
        runs: [{ text: 'Invalid' }],
      }])).toThrow(RangeError);
      expect(() => {
        shape.richText = [{ marginRight, runs: [{ text: 'Invalid' }] }];
      }).toThrow(RangeError);
    }

    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(slide.shapes[0]).toBe(shape);
    expect(slide.shapes).toHaveLength(3);
  });

  it('creates, edits, duplicates, and reopens signed ordinary paragraph indents', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const omitted = slide.addText('Canonical zero');
    const plain = slide.addText('Hanging first\r\n\rSecond', {
      align: 'center',
      paragraphIndent: -12.5,
      paragraphMarginLeft: 48,
      paragraphMarginRight: 6,
      rtlMode: false,
      spacing: { after: 4 },
      tabStops: [{ position: 1.5 }],
    });
    const rich = slide.addRichText([
      { runs: [{ text: 'Default twenty-four' }] },
      { indent: 12, runs: [{ text: 'First-line twelve' }] },
      { indent: -18, runs: [{ text: 'Hanging eighteen' }] },
      { indent: 0, runs: [{ text: 'Direct zero' }] },
      { indent: false, runs: [] },
      { bullet: true, indent: false, runs: [{ text: 'Bullet-owned indent' }] },
    ], { paragraphIndent: 24 });
    const ordinaryUnderOuterBullet = slide.addRichText([{
      bullet: false,
      indent: -18,
      runs: [{ text: 'Ordinary under outer bullet' }],
    }], { bullet: true });

    expect(omitted.richText.map(({ indent }) => indent)).toEqual([0]);
    expect(plain.richText.map(({ indent, marginLeft, marginRight }) => ({ indent, marginLeft, marginRight })))
      .toEqual([
        { indent: -12.5, marginLeft: 48, marginRight: 6 },
        { indent: -12.5, marginLeft: 48, marginRight: 6 },
        { indent: -12.5, marginLeft: 48, marginRight: 6 },
      ]);
    expect(rich.richText.map(({ indent }) => indent)).toEqual([24, 12, -18, 0, undefined, undefined]);
    expect(rich.richText[5]!.bullet).toEqual({ kind: 'bullet', character: '•', indent: 27 });
    expect(ordinaryUnderOuterBullet.richText[0]!.indent).toBe(-18);
    expect(ordinaryUnderOuterBullet.richText[0]!.bullet).toBeUndefined();
    let slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml.match(/indent="-158750"/g)).toHaveLength(3);
    expect(slideXml).toContain('indent="304800"');
    expect(slideXml).toContain('indent="152400"');
    expect(slideXml).toContain('indent="-228600"');
    expect(slideXml).toContain('indent="-342900" marL="342900"');

    document.duplicateSlide(0);
    plain.text = 'Updated\nCloned indent';
    expect(plain.richText.map(({ indent }) => indent)).toEqual([-12.5, -12.5]);
    rich.richText = [
      { indent: 6, runs: [{ text: 'Positive six' }] },
      { indent: -6, runs: [{ text: 'Negative six' }] },
      { indent: 0, runs: [{ text: 'Zero' }] },
      { indent: false, runs: [{ text: 'Cleared false' }] },
      { runs: [{ text: 'Cleared omitted' }] },
      { runs: [{ text: 'Former bullet cleared' }] },
    ];
    expect(rich.richText.map(({ indent }) => indent)).toEqual([6, -6, 0, undefined, undefined, undefined]);
    slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml).toContain('indent="76200"');
    expect(slideXml).toContain('indent="-76200"');
    const formerBulletStart = slideXml.indexOf('>Former bullet cleared<');
    const formerBulletEnd = slideXml.indexOf('</a:p>', formerBulletStart);
    const formerBullet = slideXml.slice(slideXml.lastIndexOf('<a:p>', formerBulletStart), formerBulletEnd);
    expect(formerBullet).not.toContain('indent=');
    expect(formerBullet).not.toContain('marL=');

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    expect(() =>
      document.transaction(() => {
        rich.richText = [{ indent: -48, runs: [{ text: 'Rollback' }] }];
        throw new Error('restore paragraph indent');
      }),
    ).toThrow('restore paragraph indent');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(slide.shapes[2]).toBe(rich);

    const reopened = await PptxDocument.open(await document.write());
    const editedPlain = reopened.slides[0]!.shapes[1] as ShapeModel;
    const editedRich = reopened.slides[0]!.shapes[2] as ShapeModel;
    const duplicatedRich = reopened.slides[1]!.shapes[2] as ShapeModel;
    expect(editedPlain.richText.map(({ indent }) => indent)).toEqual([-12.5, -12.5]);
    expect(editedRich.richText.map(({ indent }) => indent)).toEqual([6, -6, 0, undefined, undefined, undefined]);
    expect(duplicatedRich.richText.map(({ indent }) => indent)).toEqual([24, 12, -18, 0, undefined, undefined]);
    expect(duplicatedRich.richText[5]!.bullet).toEqual({ kind: 'bullet', character: '•', indent: 27 });
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('rejects malformed or conflicting paragraph indents before package mutation', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addText('Original');
    const positiveBoundary = slide.addText('Positive boundary', { paragraphIndent: 4032 });
    const negativeBoundary = slide.addText('Negative boundary', { paragraphIndent: -4032 });
    slide.addRichText([{ bullet: false, indent: -18, runs: [{ text: 'Ordinary' }] }], { bullet: true });
    slide.addRichText([{ bullet: true, indent: false, runs: [{ text: 'Bullet' }] }], {
      paragraphIndent: 24,
    });
    expect(positiveBoundary.richText[0]!.indent).toBe(4032);
    expect(negativeBoundary.richText[0]!.indent).toBe(-4032);
    const boundaryXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(boundaryXml).toContain('indent="51206400"');
    expect(boundaryXml).toContain('indent="-51206400"');

    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    const invalidOuterTypes = [null, true, false, '12', {}, [], Symbol('indent'), Number.NaN, Infinity];
    const invalidParagraphTypes = [null, true, '12', {}, [], Symbol('indent'), Number.NaN, Infinity];
    const invalidRanges = [-4032.01, 4032.01];

    for (const paragraphIndent of invalidOuterTypes) {
      expect(() => slide.addText('Invalid', { paragraphIndent: paragraphIndent as never })).toThrow(TypeError);
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid' }] }], {
        paragraphIndent: paragraphIndent as never,
      })).toThrow(TypeError);
    }
    for (const indent of invalidParagraphTypes) {
      expect(() => slide.addRichText([{ indent: indent as never, runs: [{ text: 'Invalid' }] }]))
        .toThrow(TypeError);
      expect(() => {
        shape.richText = [{ indent: indent as never, runs: [{ text: 'Invalid' }] }];
      }).toThrow(TypeError);
    }
    for (const indent of invalidRanges) {
      expect(() => slide.addText('Invalid', { paragraphIndent: indent })).toThrow(RangeError);
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid' }] }], { paragraphIndent: indent }))
        .toThrow(RangeError);
      expect(() => slide.addRichText([{ indent, runs: [{ text: 'Invalid' }] }])).toThrow(RangeError);
      expect(() => {
        shape.richText = [{ indent, runs: [{ text: 'Invalid' }] }];
      }).toThrow(RangeError);
    }
    expect(() => slide.addText('Conflict', { bullet: true, paragraphIndent: 0 })).toThrow(TypeError);
    expect(() => slide.addRichText([{ runs: [{ text: 'Conflict' }] }], {
      bullet: true,
      paragraphIndent: 12,
    })).toThrow(TypeError);
    expect(() => slide.addRichText([{
      bullet: true,
      indent: 0,
      runs: [{ text: 'Conflict' }],
    }])).toThrow(TypeError);
    expect(() => {
      shape.richText = [{ bullet: true, indent: -12, runs: [{ text: 'Conflict' }] }];
    }).toThrow(TypeError);

    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(slide.shapes[0]).toBe(shape);
    expect(slide.shapes).toHaveLength(5);
  });

  it('creates, replaces, rolls back, and round-trips bullets and numbering', async () => {
    const styles: readonly NumberingStyle[] = [
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
    ];
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const plain = slide.addText('First\nSecond', { bullet: true });
    const rich = slide.addRichText(
      [
        { runs: [{ text: 'Default' }] },
        { runs: [{ text: 'Custom' }], bullet: { kind: 'bullet', character: '▶', indent: 18 } },
        {
          runs: [{ text: 'Roman' }],
          bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
        },
        { runs: [], bullet: { kind: 'bullet', character: '💡', indent: 21 } },
        { runs: [], bullet: false },
      ],
      { bullet: true },
    );
    const allStyles = slide.addRichText(
      styles.map((style, index) => ({
        runs: [{ text: style }],
        bullet: { kind: 'number' as const, style, startAt: index + 1, indent: 20 },
      })),
    );

    expect(plain.richText.map(({ bullet }) => bullet)).toEqual([
      { kind: 'bullet', character: '•', indent: 27 },
      { kind: 'bullet', character: '•', indent: 27 },
    ]);
    expect(rich.richText.map(({ bullet }) => bullet)).toEqual([
      { kind: 'bullet', character: '•', indent: 27 },
      { kind: 'bullet', character: '▶', indent: 18 },
      { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
      { kind: 'bullet', character: '💡', indent: 21 },
      undefined,
    ]);
    expect(allStyles.richText.map(({ bullet }) =>
      bullet && typeof bullet !== 'boolean' && bullet.kind === 'number' ? bullet.style : undefined,
    )).toEqual(styles);
    let slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml).toContain('<a:buChar char="▶"/>');
    expect(slideXml).toContain('<a:buChar char="💡"/>');
    expect(slideXml).toContain('<a:buAutoNum type="romanUcPeriod" startAt="3"/>');
    expect(slideXml).toContain('indent="-279400" marL="279400"');
    const bulletSnapshot = rich.richText as unknown as Array<{ bullet?: { indent: number } }>;
    bulletSnapshot[0]!.bullet!.indent = 999;
    expect(rich.richText[0]!.bullet).toEqual({ kind: 'bullet', character: '•', indent: 27 });

    rich.richText = [
      {
        runs: rich.richText[0]!.runs,
        bullet: { kind: 'number', style: 'romanLcParenR', startAt: 4, indent: 24 },
      },
      { runs: rich.richText[1]!.runs },
      { runs: rich.richText[2]!.runs, bullet: true },
      { runs: rich.richText[3]!.runs, bullet: false },
      { runs: rich.richText[4]!.runs, bullet: false },
    ];
    expect(rich.richText.map(({ bullet }) => bullet)).toEqual([
      { kind: 'number', style: 'romanLcParenR', startAt: 4, indent: 24 },
      undefined,
      { kind: 'bullet', character: '•', indent: 27 },
      undefined,
      undefined,
    ]);
    slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml).not.toContain('<a:buChar char="▶"/>');

    expect(() =>
      document.transaction(() => {
        rich.richText = [{ runs: [{ text: 'Rollback' }], bullet: { kind: 'bullet', character: '◆' } }];
        throw new Error('restore paragraph bullets');
      }),
    ).toThrow('restore paragraph bullets');
    expect(rich.richText[0]!.bullet).toEqual({
      kind: 'number',
      style: 'romanLcParenR',
      startAt: 4,
      indent: 24,
    });

    const reopened = await PptxDocument.open(await document.write());
    const reopenedPlain = reopened.slides[0]!.shapes[0] as ShapeModel;
    const reopenedRich = reopened.slides[0]!.shapes[1] as ShapeModel;
    expect(reopenedPlain.richText[0]!.bullet).toEqual({ kind: 'bullet', character: '•', indent: 27 });
    expect(reopenedRich.richText.map(({ bullet }) => bullet)).toEqual(rich.richText.map(({ bullet }) => bullet));
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('creates and reopens bullet and numbering owners in all six formats', async () => {
    const styles: readonly NumberingStyle[] = [
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
    ];
    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({ format });
      const slide = created.addSlide();
      const allStyles = slide.addRichText(styles.map((style, index) => ({
        runs: [{ text: style }],
        bullet: { kind: 'number' as const, style, startAt: index + 1, indent: 20 },
      })), { name: 'all-numbering-styles' });
      slide.addText('Default\nCustom', { bullet: true, name: 'plain-bullets' });
      allStyles.richText = [
        ...allStyles.richText,
        { runs: [{ text: 'Custom' }], bullet: { kind: 'bullet', character: '▶', indent: 18 } },
      ];
      created.layouts[0]!.addPlaceholder('Placeholder number', {
        name: 'numbered-placeholder',
        type: 'body',
        index: 100,
        bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
      });
      slide.addTable([[{
        text: [{
          runs: [{ text: 'Table cell bullet' }],
          bullet: { kind: 'bullet', character: '◆', indent: 19 },
        }],
      }]], { name: 'bullet-table' });

      const first = await PptxDocument.open(await created.write());
      const firstStyles = first.slides[0]!.shapes.find(
        (shape): shape is ShapeModel => shape instanceof ShapeModel
          && shape.name === 'all-numbering-styles',
      )!;
      expect(firstStyles.richText.slice(0, styles.length).map(({ bullet }) =>
        bullet && typeof bullet !== 'boolean' && bullet.kind === 'number'
          ? bullet.style
          : undefined,
      )).toEqual(styles);
      expect(firstStyles.richText.at(-1)?.bullet).toEqual({
        kind: 'bullet', character: '▶', indent: 18,
      });
      const firstPlaceholder = first.layouts[0]!.placeholders.find(
        (shape): shape is ShapeModel => shape instanceof ShapeModel
          && shape.name === 'numbered-placeholder',
      )!;
      expect(firstPlaceholder.richText[0]?.bullet).toEqual({
        kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22,
      });
      const firstTable = first.slides[0]!.shapes.find(
        (shape): shape is TableModel => shape instanceof TableModel && shape.name === 'bullet-table',
      )!;
      expect(firstTable.rows[0]!.cells[0]!.richText[0]?.bullet).toEqual({
        kind: 'bullet', character: '◆', indent: 19,
      });
      const slidePart = new TextDecoder().decode(
        first.opcPackage.requirePart(first.slides[0]!.partUri).bytes,
      );
      for (const [index, style] of styles.entries()) {
        expect(slidePart).toContain(
          `<a:buAutoNum type="${style}" startAt="${index + 1}"/>`,
        );
      }
      expect(slidePart).toMatch(/<a:pPr\b[^>]*\bindent="-254000"[^>]*\bmarL="254000"/u);
      expect(slidePart).toContain('<a:buChar char="◆"/>');

      firstStyles.richText = [{ runs: [{ text: 'Cleared text bullet' }], bullet: false }];
      firstPlaceholder.richText = [{
        runs: [{ text: 'Edited placeholder bullet' }],
        bullet: { kind: 'bullet', character: '•', indent: 27 },
      }];
      firstTable.setCellRichText(0, 0, [{
        runs: [{ text: 'Edited table number' }],
        bullet: { kind: 'number', style: 'alphaUcPeriod', startAt: 2, indent: 21 },
      }]);
      const second = await PptxDocument.open(await first.write());
      expect((second.slides[0]!.shapes.find(
        (shape): shape is ShapeModel => shape instanceof ShapeModel
          && shape.name === 'all-numbering-styles',
      )!).richText[0]?.bullet).toBeUndefined();
      expect((second.layouts[0]!.placeholders.find(
        (shape): shape is ShapeModel => shape instanceof ShapeModel
          && shape.name === 'numbered-placeholder',
      )!).richText[0]?.bullet).toEqual({
        kind: 'bullet', character: '•', indent: 27,
      });
      expect((second.slides[0]!.shapes.find(
        (shape): shape is TableModel => shape instanceof TableModel && shape.name === 'bullet-table',
      )!).rows[0]!.cells[0]!.richText[0]?.bullet).toEqual({
        kind: 'number', style: 'alphaUcPeriod', startAt: 2, indent: 21,
      });
      expect(validatePackage(second.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);
    }
  });

  it('creates edits clears and reopens tab stops in all six formats', async () => {
    const tabStops = [
      { position: 1.25, alignment: 'left' as const },
      { position: 2.5, alignment: 'center' as const },
      { position: 3.75, alignment: 'right' as const },
      { position: 4.5, alignment: 'decimal' as const },
    ];
    const expectedTabXml = '<a:tab pos="1143000" algn="l"/>'
      + '<a:tab pos="2286000" algn="ctr"/>'
      + '<a:tab pos="3429000" algn="r"/>'
      + '<a:tab pos="4114800" algn="dec"/>';
    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({ format });
      const layout = created.layouts[0]!;
      layout.addPlaceholder('Placeholder\t100', {
        name: 'tab-stops-placeholder',
        type: 'body',
        index: 101,
        tabStops,
      });
      const slide = created.addSlide();
      slide.addText('Text\t12.50', { name: 'tab-stops-text', tabStops });
      slide.addTable([[{
        text: [{ runs: [{ text: 'Cell\t42.75' }], tabStops }],
      }]], { name: 'tab-stops-table' });

      const beforeInvalidBytes = created.opcPackage.requirePart(slide.partUri).bytes.slice();
      const beforeInvalidJournal = [...created.opcPackage.mutations];
      const beforeInvalidShapeCount = slide.shapes.length;
      expect(() => slide.addText('Invalid tab stop', {
        tabStops: [{ position: 1, alignment: 'tab' as never }],
      })).toThrow(/alignment/u);
      expect(created.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeInvalidBytes);
      expect(created.opcPackage.mutations).toEqual(beforeInvalidJournal);
      expect(slide.shapes).toHaveLength(beforeInvalidShapeCount);
      expect(created.diagnostics).toEqual([]);

      const first = await PptxDocument.open(await created.write());
      expect(first.diagnostics).toEqual([]);
      const firstText = first.slides[0]!.shapes.find(
        (shape): shape is ShapeModel => shape instanceof ShapeModel
          && shape.name === 'tab-stops-text',
      )!;
      const firstPlaceholder = first.layouts[0]!.placeholders.find(
        (shape): shape is ShapeModel => shape instanceof ShapeModel
          && shape.name === 'tab-stops-placeholder',
      )!;
      const firstTable = first.slides[0]!.shapes.find(
        (shape): shape is TableModel => shape instanceof TableModel
          && shape.name === 'tab-stops-table',
      )!;
      expect(firstText.richText[0]!.tabStops).toEqual(tabStops);
      expect(firstPlaceholder.richText[0]!.tabStops).toEqual(tabStops);
      expect(firstTable.rows[0]!.cells[0]!.richText[0]!.tabStops).toEqual(tabStops);

      const detached = firstText.richText as unknown as Array<{
        tabStops?: Array<{ position: number }>;
      }>;
      detached[0]!.tabStops![0]!.position = 99;
      expect(firstText.richText[0]!.tabStops).toEqual(tabStops);

      const firstSlideXml = new TextDecoder().decode(
        first.opcPackage.requirePart(first.slides[0]!.partUri).bytes,
      );
      const firstLayoutXml = new TextDecoder().decode(
        first.opcPackage.requirePart(first.layouts[0]!.partUri).bytes,
      );
      expect((firstSlideXml.match(new RegExp(expectedTabXml, 'g')) ?? []).length).toBe(2);
      expect(firstLayoutXml).toContain(expectedTabXml);

      firstText.richText = [{
        runs: [{ text: 'Edited\t12.50' }],
        tabStops: [{ position: 2.75, alignment: 'decimal' }],
      }];
      firstPlaceholder.richText = [{ runs: [{ text: 'Explicit empty' }], tabStops: [] }];
      firstTable.setCellRichText(0, 0, [{ runs: [{ text: 'Cleared' }], tabStops: false }]);
      const second = await PptxDocument.open(await first.write());
      expect(second.diagnostics).toEqual([]);
      const secondText = second.slides[0]!.shapes.find(
        (shape): shape is ShapeModel => shape instanceof ShapeModel
          && shape.name === 'tab-stops-text',
      )!;
      const secondPlaceholder = second.layouts[0]!.placeholders.find(
        (shape): shape is ShapeModel => shape instanceof ShapeModel
          && shape.name === 'tab-stops-placeholder',
      )!;
      const secondTable = second.slides[0]!.shapes.find(
        (shape): shape is TableModel => shape instanceof TableModel
          && shape.name === 'tab-stops-table',
      )!;
      expect(secondText.richText[0]!.tabStops).toEqual([
        { position: 2.75, alignment: 'decimal' },
      ]);
      expect(secondPlaceholder.richText[0]!.tabStops).toEqual([]);
      expect(secondTable.rows[0]!.cells[0]!.richText[0]!.tabStops).toBeUndefined();
      const secondSlideXml = new TextDecoder().decode(
        second.opcPackage.requirePart(second.slides[0]!.partUri).bytes,
      );
      const secondLayoutXml = new TextDecoder().decode(
        second.opcPackage.requirePart(second.layouts[0]!.partUri).bytes,
      );
      expect(secondSlideXml).toContain('<a:tab pos="2514600" algn="dec"/>');
      expect((secondSlideXml.match(/<a:tab\b/g) ?? []).length).toBe(1);
      expect(secondLayoutXml).toContain('<a:tabLst></a:tabLst>');
      expect(validatePackage(second.opcPackage).filter(({ severity }) => severity === 'error'))
        .toEqual([]);
    }
  });

  it('creates, replaces, rolls back, and round-trips paragraph spacing', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const plain = slide.addText('Exact first\nExact second', {
      spacing: { before: 6.25, after: 8.5, line: { kind: 'exact', points: 28 } },
    });
    const rich = slide.addRichText(
      [
        { runs: [{ text: 'Defaults' }], bullet: true },
        {
          runs: [{ text: 'Partial override' }],
          spacing: { before: 0, line: { kind: 'multiple', factor: 1.5 } },
        },
        { runs: [], spacing: { after: 12, line: false } },
        { runs: [{ text: 'No spacing' }], spacing: false },
      ],
      { spacing: { before: 4, after: 6, line: { kind: 'multiple', factor: 1.2 } } },
    );

    expect(plain.richText.map(({ spacing }) => spacing)).toEqual([
      { before: 6.25, after: 8.5, line: { kind: 'exact', points: 28 } },
      { before: 6.25, after: 8.5, line: { kind: 'exact', points: 28 } },
    ]);
    expect(rich.richText.map(({ spacing }) => spacing)).toEqual([
      { before: 4, after: 6, line: { kind: 'multiple', factor: 1.2 } },
      { after: 6, line: { kind: 'multiple', factor: 1.5 } },
      { before: 4, after: 12 },
      undefined,
    ]);
    let slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml).toContain('<a:lnSpc><a:spcPts val="2800"/></a:lnSpc>');
    expect(slideXml).toContain('<a:lnSpc><a:spcPct val="150000"/></a:lnSpc>');
    expect(slideXml).toContain('<a:spcBef><a:spcPts val="625"/></a:spcBef>');
    const firstRichProperties = slideXml.slice(
      slideXml.indexOf('<a:pPr', slideXml.indexOf('Defaults') - 500),
      slideXml.indexOf('</a:pPr>', slideXml.indexOf('Defaults') - 500),
    );
    expect(firstRichProperties.indexOf('<a:lnSpc>')).toBeLessThan(firstRichProperties.indexOf('<a:spcBef>'));
    expect(firstRichProperties.indexOf('<a:spcBef>')).toBeLessThan(firstRichProperties.indexOf('<a:spcAft>'));
    expect(firstRichProperties.indexOf('<a:spcAft>')).toBeLessThan(firstRichProperties.indexOf('<a:buSzPct'));
    const spacingSnapshot = rich.richText as unknown as Array<{ spacing?: { before?: number } }>;
    spacingSnapshot[0]!.spacing!.before = 999;
    expect(rich.richText[0]!.spacing).toEqual({
      before: 4,
      after: 6,
      line: { kind: 'multiple', factor: 1.2 },
    });

    rich.richText = [
      {
        runs: rich.richText[0]!.runs,
        bullet: rich.richText[0]!.bullet!,
        spacing: { before: 5, line: { kind: 'exact', points: 20 } },
      },
      { runs: rich.richText[1]!.runs, spacing: false },
      { runs: rich.richText[2]!.runs, spacing: { after: 7.5 } },
      { runs: rich.richText[3]!.runs },
    ];
    expect(rich.richText.map(({ spacing }) => spacing)).toEqual([
      { before: 5, line: { kind: 'exact', points: 20 } },
      undefined,
      { after: 7.5 },
      undefined,
    ]);
    slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml).toContain('<a:spcAft><a:spcPts val="750"/></a:spcAft>');
    expect(slideXml).not.toContain('val="150000"');

    expect(() =>
      document.transaction(() => {
        rich.richText = [{
          runs: [{ text: 'Rollback' }],
          spacing: { line: { kind: 'multiple', factor: 2 } },
        }];
        throw new Error('restore paragraph spacing');
      }),
    ).toThrow('restore paragraph spacing');
    expect(rich.richText[0]!.spacing).toEqual({
      before: 5,
      line: { kind: 'exact', points: 20 },
    });

    const reopened = await PptxDocument.open(await document.write());
    const reopenedPlain = reopened.slides[0]!.shapes[0] as ShapeModel;
    const reopenedRich = reopened.slides[0]!.shapes[1] as ShapeModel;
    expect(reopenedPlain.richText.map(({ spacing }) => spacing)).toEqual(
      plain.richText.map(({ spacing }) => spacing),
    );
    expect(reopenedRich.richText.map(({ spacing }) => spacing)).toEqual(
      rich.richText.map(({ spacing }) => spacing),
    );
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('creates, replaces, rolls back, and round-trips paragraph list levels', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const plain = slide.addText('Nested first\nNested second', { bullet: true, level: 2 });
    const rich = slide.addRichText(
      [
        { runs: [{ text: 'Default level' }], spacing: { after: 4 } },
        {
          runs: [{ text: 'Custom level two' }],
          bullet: { kind: 'bullet', character: '▶', indent: 18 },
          level: 2,
        },
        {
          runs: [{ text: 'Root number' }],
          bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
          level: 0,
        },
        { runs: [], bullet: false, level: 3 },
      ],
      { bullet: true, level: 1 },
    );
    const allLevels = slide.addRichText(
      Array.from({ length: 9 }, (_, level) => ({
        runs: [{ text: `Level ${level}` }],
        bullet: { kind: 'bullet' as const, character: '◆', indent: 18 },
        level,
      })),
    );

    expect(plain.richText.map(({ bullet, level }) => ({ bullet, level }))).toEqual([
      { bullet: { kind: 'bullet', character: '•', indent: 27 }, level: 2 },
      { bullet: { kind: 'bullet', character: '•', indent: 27 }, level: 2 },
    ]);
    expect(rich.richText.map(({ bullet, level }) => ({ bullet, level }))).toEqual([
      { bullet: { kind: 'bullet', character: '•', indent: 27 }, level: 1 },
      { bullet: { kind: 'bullet', character: '▶', indent: 18 }, level: 2 },
      {
        bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
        level: undefined,
      },
      { bullet: undefined, level: 3 },
    ]);
    expect(allLevels.richText.map(({ bullet, level }) => ({
      indent: bullet && typeof bullet !== 'boolean' ? bullet.indent : undefined,
      level,
    }))).toEqual(Array.from({ length: 9 }, (_, level) => ({
      indent: 18,
      level: level === 0 ? undefined : level,
    })));
    let slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml).toContain('indent="-342900" marL="1028700" lvl="2"');
    expect(slideXml).toContain('indent="-228600" marL="685800" lvl="2"');
    expect(slideXml).toContain('indent="-228600" marL="2057400" lvl="8"');
    expect(slideXml).toContain('<a:spcAft><a:spcPts val="400"/></a:spcAft><a:buSzPct');

    rich.richText = [
      {
        runs: rich.richText[0]!.runs,
        bullet: rich.richText[0]!.bullet!,
        level: 3,
        spacing: rich.richText[0]!.spacing!,
      },
      { runs: rich.richText[1]!.runs, bullet: rich.richText[1]!.bullet!, level: 0 },
      { runs: rich.richText[2]!.runs, bullet: rich.richText[2]!.bullet!, level: 2 },
      { runs: rich.richText[3]!.runs, bullet: false, level: 0 },
    ];
    expect(rich.richText.map(({ bullet, level }) => ({ bullet, level }))).toEqual([
      { bullet: { kind: 'bullet', character: '•', indent: 27 }, level: 3 },
      { bullet: { kind: 'bullet', character: '▶', indent: 18 }, level: undefined },
      {
        bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
        level: 2,
      },
      { bullet: undefined, level: undefined },
    ]);
    slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml).toContain('indent="-342900" marL="1371600" lvl="3"');
    expect(slideXml).toContain('indent="-228600" marL="228600"');
    expect(slideXml).toContain('indent="-279400" marL="838200" lvl="2"');

    expect(() =>
      document.transaction(() => {
        rich.richText = [{ runs: [{ text: 'Rollback' }], bullet: true, level: 4 }];
        throw new Error('restore paragraph levels');
      }),
    ).toThrow('restore paragraph levels');
    expect(rich.richText[0]!.level).toBe(3);
    expect(() => slide.addText('Too wide', {
      bullet: { kind: 'bullet', indent: 500 },
      level: 8,
    })).toThrow(/4032 points total/);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedPlain = reopened.slides[0]!.shapes[0] as ShapeModel;
    const reopenedRich = reopened.slides[0]!.shapes[1] as ShapeModel;
    expect(reopenedPlain.richText.map(({ bullet, level }) => ({ bullet, level }))).toEqual(
      plain.richText.map(({ bullet, level }) => ({ bullet, level })),
    );
    expect(reopenedRich.richText.map(({ bullet, level }) => ({ bullet, level }))).toEqual(
      rich.richText.map(({ bullet, level }) => ({ bullet, level })),
    );
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('creates, replaces, rolls back, and round-trips paragraph tab stops', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const sourceStops = [
      { position: 1.25, alignment: 'left' as const },
      { position: 2.5, alignment: 'right' as const },
    ];
    const plain = slide.addText('Name\tValue\nCount\t12.50\n', { tabStops: sourceStops });
    const rich = slide.addRichText(
      [
        { runs: [{ text: 'Default\t100' }], spacing: { after: 4 } },
        {
          runs: [{ text: 'Override\t12.50' }],
          tabStops: [
            { position: 0, alignment: 'center' },
            { position: -0.5, alignment: 'decimal' },
          ],
        },
        { runs: [{ text: 'Suppressed' }], tabStops: false },
        { runs: [], tabStops: [] },
      ],
      { tabStops: [{ position: 1.5 }, { position: 2.75, alignment: 'right' }] },
    );
    const empty = slide.addText('Empty list', { tabStops: [] });
    sourceStops[0]!.position = 9;
    sourceStops.push({ position: 10, alignment: 'left' });

    const plainStops = [
      { position: 1.25, alignment: 'left' },
      { position: 2.5, alignment: 'right' },
    ];
    expect(plain.richText.map(({ tabStops }) => tabStops)).toEqual([
      plainStops,
      plainStops,
      plainStops,
    ]);
    expect(rich.richText.map(({ tabStops }) => tabStops)).toEqual([
      [
        { position: 1.5, alignment: 'left' },
        { position: 2.75, alignment: 'right' },
      ],
      [
        { position: 0, alignment: 'center' },
        { position: -0.5, alignment: 'decimal' },
      ],
      undefined,
      [],
    ]);
    expect(empty.richText[0]!.tabStops).toEqual([]);

    let slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml).toContain('<a:tab pos="1143000" algn="l"/><a:tab pos="2286000" algn="r"/>');
    expect(slideXml).toContain('<a:tab pos="0" algn="ctr"/><a:tab pos="-457200" algn="dec"/>');
    expect(slideXml).toContain('<a:tabLst></a:tabLst>');
    expect(slideXml).toContain('<a:buNone/><a:tabLst>');

    const snapshot = rich.richText as unknown as Array<{ tabStops?: Array<{ position: number }> }>;
    snapshot[0]!.tabStops![0]!.position = 99;
    expect(rich.richText[0]!.tabStops).toEqual([
      { position: 1.5, alignment: 'left' },
      { position: 2.75, alignment: 'right' },
    ]);

    plain.text = 'Updated\t10\nAgain\t20';
    expect(plain.richText.map(({ tabStops }) => tabStops)).toEqual([plainStops, plainStops]);

    rich.richText = [
      {
        runs: rich.richText[0]!.runs,
        spacing: rich.richText[0]!.spacing!,
        tabStops: [
          { position: 3, alignment: 'right' },
          { position: 1, alignment: 'left' },
          { position: 3, alignment: 'center' },
        ],
      },
      { runs: rich.richText[1]!.runs, tabStops: false },
      { runs: rich.richText[2]!.runs, tabStops: [] },
      { runs: rich.richText[3]!.runs },
    ];
    expect(rich.richText.map(({ tabStops }) => tabStops)).toEqual([
      [
        { position: 3, alignment: 'right' },
        { position: 1, alignment: 'left' },
        { position: 3, alignment: 'center' },
      ],
      undefined,
      [],
      undefined,
    ]);
    slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml).toContain('<a:tab pos="2743200" algn="r"/><a:tab pos="914400" algn="l"/><a:tab pos="2743200" algn="ctr"/>');

    expect(() =>
      document.transaction(() => {
        rich.richText = [{ runs: [{ text: 'Rollback' }], tabStops: [{ position: 4 }] }];
        throw new Error('restore paragraph tab stops');
      }),
    ).toThrow('restore paragraph tab stops');
    expect(rich.richText[0]!.tabStops).toEqual([
      { position: 3, alignment: 'right' },
      { position: 1, alignment: 'left' },
      { position: 3, alignment: 'center' },
    ]);
    expect(() => slide.addText('Too far', { tabStops: [{ position: 3_000 }] })).toThrow(/signed 32-bit/);

    const reopened = await PptxDocument.open(await document.write());
    const reopenedPlain = reopened.slides[0]!.shapes[0] as ShapeModel;
    const reopenedRich = reopened.slides[0]!.shapes[1] as ShapeModel;
    const reopenedEmpty = reopened.slides[0]!.shapes[2] as ShapeModel;
    expect(reopenedPlain.richText.map(({ tabStops }) => tabStops)).toEqual(
      plain.richText.map(({ tabStops }) => tabStops),
    );
    expect(reopenedRich.richText.map(({ tabStops }) => tabStops)).toEqual(
      rich.richText.map(({ tabStops }) => tabStops),
    );
    expect(reopenedEmpty.richText[0]!.tabStops).toEqual([]);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('creates, reads, replaces, and round-trips rich text run styles', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const input = [
      {
        runs: [
          {
            text: 'Bold & ',
            style: {
              fontFamily: 'Aptos & Display',
              fontSize: 12.5,
              bold: true,
              italic: false,
              color: { kind: 'srgb' as const, value: '#ff0000' },
            },
          },
          {
            text: 'soft',
            softBreakBefore: true,
            style: { italic: true, color: { kind: 'scheme' as const, value: 'accent1' } },
          },
        ],
      },
      { runs: [] },
      { runs: [{ text: 'Last' }] },
    ];
    const shape = slide.addRichText(input, {
      name: 'Rich text',
      x: inches(1),
      y: inches(1),
      width: inches(5),
      height: inches(2),
    });
    input[0]!.runs[0]!.text = 'MUTATED';

    expect(shape.text).toBe('Bold & \nsoft\n\nLast');
    expect(shape.richText).toEqual([
      {
        indent: 0,
        marginLeft: 0,
        runs: [
          {
            text: 'Bold & ',
            style: {
              fontFamily: 'Aptos & Display',
              fontSize: 12.5,
              lang: 'en-US',
              bold: true,
              italic: false,
              color: { kind: 'srgb', value: 'FF0000' },
            },
          },
          {
            text: 'soft',
            softBreakBefore: true,
            style: {
              fontFamily: '+mn-lt',
              lang: 'en-US',
              italic: true,
              color: { kind: 'scheme', value: 'accent1' },
            },
          },
        ],
      },
      { indent: 0, marginLeft: 0, runs: [] },
      {
        indent: 0,
        marginLeft: 0,
        runs: [
          {
            text: 'Last',
            style: { fontFamily: '+mn-lt', lang: 'en-US', color: { kind: 'scheme', value: 'tx1' } },
          },
        ],
      },
    ]);
    const createdXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(createdXml).toContain('typeface="Aptos &amp; Display"');
    expect(createdXml).toContain('sz="1250" b="1" i="0"');
    expect(createdXml).toContain('<a:srgbClr val="FF0000"/>');
    expect(createdXml).toContain('<a:ea typeface="+mn-ea"/>');
    expect(createdXml).toContain('<a:cs typeface="+mn-cs"/>');
    const snapshot = shape.richText as Array<{ runs: Array<{ text: string }> }>;
    snapshot[0]!.runs[0]!.text = 'LOCAL';
    expect(shape.text).toBe('Bold & \nsoft\n\nLast');

    shape.richText = [
      { runs: [{ text: 'Updated', style: { bold: false, fontSize: 24 } }] },
      { runs: [{ text: 'Blue', style: { color: { kind: 'scheme', value: 'tx2' } } }] },
    ];
    expect(shape.text).toBe('Updated\nBlue');
    expect(slide.shapes[0]).toBe(shape);
    expect(shape.richText[0]!.runs[0]).toMatchObject({ text: 'Updated', style: { bold: false, fontSize: 24 } });

    const slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml).toContain('sz="2400" b="0"');
    expect(slideXml).toContain('<a:schemeClr val="tx2"/>');

    expect(() =>
      document.transaction(() => {
        shape.richText = [{ runs: [{ text: 'Rollback', style: { italic: true } }] }];
        throw new Error('restore rich text runs');
      }),
    ).toThrow('restore rich text runs');
    expect(shape.text).toBe('Updated\nBlue');

    const reopened = await PptxDocument.open(await document.write());
    const reopenedShape = reopened.slides[0]!.shapes[0] as ShapeModel;
    expect(reopenedShape.text).toBe('Updated\nBlue');
    expect(reopenedShape.richText[0]!.runs[0]).toMatchObject({ style: { bold: false, fontSize: 24 } });
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('creates, edits, duplicates, and reopens rich text outlines', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const sourceOutline = {
      color: { kind: 'srgb' as const, value: '#ff0000' },
      size: 1.5,
    };
    const shape = slide.addRichText([{
      runs: [
        { text: 'Red', style: { outline: sourceOutline } },
        { text: 'Theme', style: { outline: { color: { kind: 'scheme', value: 'accent1' }, size: 2 } } },
        { text: 'Hairline', style: { outline: { color: { kind: 'srgb', value: '00FF00' }, size: 0 } } },
        { text: 'Quantized', style: { outline: { color: { kind: 'srgb', value: '0000FF' }, size: 1.23456 } } },
        { text: 'Maximum', style: { outline: { color: { kind: 'scheme', value: 'tx2' }, size: 1584 } } },
        {
          text: 'Combined',
          style: {
            highlight: { kind: 'srgb', value: 'FFFF00' },
            outline: { color: { kind: 'scheme', value: 'accent2' }, size: 0.75 },
            strike: 'dblStrike',
            underline: { style: 'dbl' },
          },
        },
      ],
    }]);
    sourceOutline.color.value = '000000';
    sourceOutline.size = 3;

    expect(shape.richText[0]!.runs.map(({ style }) => style?.outline)).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, size: 1.5 },
      { color: { kind: 'scheme', value: 'accent1' }, size: 2 },
      { color: { kind: 'srgb', value: '00FF00' }, size: 0 },
      { color: { kind: 'srgb', value: '0000FF' }, size: 15_679 / 12_700 },
      { color: { kind: 'scheme', value: 'tx2' }, size: 1584 },
      { color: { kind: 'scheme', value: 'accent2' }, size: 0.75 },
    ]);
    const createdXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(createdXml).toContain(
      '<a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln><a:solidFill>',
    );
    expect(createdXml).toContain('<a:ln w="15679">');
    expect(createdXml).toContain('<a:ln w="20116800">');

    const snapshot = shape.richText as unknown as Array<{
      runs: Array<{ style?: { outline?: { color: { value: string }; size: number } } }>;
    }>;
    snapshot[0]!.runs[0]!.style!.outline!.color.value = 'FFFFFF';
    snapshot[0]!.runs[0]!.style!.outline!.size = 4;
    expect(shape.richText[0]!.runs[0]!.style!.outline).toEqual({
      color: { kind: 'srgb', value: 'FF0000' },
      size: 1.5,
    });

    document.duplicateSlide(0);
    shape.richText = [{
      runs: [
        { text: 'Changed', style: { outline: { color: { kind: 'scheme', value: 'tx1' }, size: 1 } } },
        { text: 'Cleared' },
      ],
    }];
    expect(shape.richText[0]!.runs.map(({ style }) => style?.outline)).toEqual([
      { color: { kind: 'scheme', value: 'tx1' }, size: 1 },
      undefined,
    ]);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    expect(() =>
      document.transaction(() => {
        shape.richText = [{
          runs: [{
            text: 'Rollback',
            style: { outline: { color: { kind: 'srgb', value: '000000' }, size: 2 } },
          }],
        }];
        throw new Error('restore outline');
      }),
    ).toThrow('restore outline');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(slide.shapes[0]).toBe(shape);

    const reopened = await PptxDocument.open(await document.write());
    const edited = reopened.slides[0]!.shapes[0] as ShapeModel;
    const duplicated = reopened.slides[1]!.shapes[0] as ShapeModel;
    expect(edited.richText[0]!.runs.map(({ style }) => style?.outline)).toEqual([
      { color: { kind: 'scheme', value: 'tx1' }, size: 1 },
      undefined,
    ]);
    expect(duplicated.richText[0]!.runs[0]!.style!.outline).toEqual({
      color: { kind: 'srgb', value: 'FF0000' },
      size: 1.5,
    });
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('creates, edits, duplicates, and reopens rich text glows', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const sourceGlow = {
      color: { kind: 'srgb' as const, value: '#ff0000' },
      opacity: 0.5,
      size: 8,
    };
    const shape = slide.addRichText([{
      runs: [
        { text: 'Red', style: { glow: sourceGlow } },
        { text: 'Default', style: { glow: { opacity: 0.75, size: 4 } } },
        { text: 'Theme', style: { glow: { color: { kind: 'scheme', value: 'accent1' }, opacity: 1, size: 2.5 } } },
        { text: 'Zero', style: { glow: { color: { kind: 'srgb', value: '00FF00' }, opacity: 0, size: 0 } } },
        { text: 'Quantized', style: { glow: { color: { kind: 'srgb', value: '0000FF' }, opacity: 0.123456, size: 1.23456 } } },
        { text: 'Maximum', style: { glow: { color: { kind: 'scheme', value: 'accent2' }, opacity: 1, size: 2_147_483_647 } } },
        {
          text: 'Combined',
          style: {
            glow: { color: { kind: 'scheme', value: 'accent3' }, opacity: 0.25, size: 6 },
            highlight: { kind: 'srgb', value: 'FFFF00' },
            outline: { color: { kind: 'scheme', value: 'accent2' }, size: 0.75 },
            strike: 'dblStrike',
            underline: { style: 'dbl' },
          },
        },
      ],
    }]);
    sourceGlow.color.value = '000000';
    sourceGlow.opacity = 1;
    sourceGlow.size = 3;

    expect(shape.richText[0]!.runs.map(({ style }) => style?.glow)).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, opacity: 0.5, size: 8 },
      { color: { kind: 'srgb', value: 'FFFFFF' }, opacity: 0.75, size: 4 },
      { color: { kind: 'scheme', value: 'accent1' }, opacity: 1, size: 2.5 },
      { color: { kind: 'srgb', value: '00FF00' }, opacity: 0, size: 0 },
      { color: { kind: 'srgb', value: '0000FF' }, opacity: 0.12346, size: 15_679 / 12_700 },
      { color: { kind: 'scheme', value: 'accent2' }, opacity: 1, size: 2_147_483_647 },
      { color: { kind: 'scheme', value: 'accent3' }, opacity: 0.25, size: 6 },
    ]);
    const createdXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(createdXml).toContain(
      '<a:effectLst><a:glow rad="101600"><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:glow></a:effectLst>',
    );
    expect(createdXml).toContain('<a:glow rad="15679"><a:srgbClr val="0000FF"><a:alpha val="12346"/>');
    expect(createdXml).toContain('<a:glow rad="27273042316900">');

    const snapshot = shape.richText as unknown as Array<{
      runs: Array<{ style?: { glow?: { color?: { value: string }; opacity: number; size: number } } }>;
    }>;
    snapshot[0]!.runs[0]!.style!.glow!.color!.value = 'FFFFFF';
    snapshot[0]!.runs[0]!.style!.glow!.opacity = 1;
    snapshot[0]!.runs[0]!.style!.glow!.size = 4;
    expect(shape.richText[0]!.runs[0]!.style!.glow).toEqual({
      color: { kind: 'srgb', value: 'FF0000' },
      opacity: 0.5,
      size: 8,
    });

    document.duplicateSlide(0);
    shape.richText = [{
      runs: [
        { text: 'Changed', style: { glow: { color: { kind: 'scheme', value: 'tx1' }, opacity: 0.6, size: 5 } } },
        { text: 'Cleared' },
      ],
    }];
    expect(shape.richText[0]!.runs.map(({ style }) => style?.glow)).toEqual([
      { color: { kind: 'scheme', value: 'tx1' }, opacity: 0.6, size: 5 },
      undefined,
    ]);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    expect(() =>
      document.transaction(() => {
        shape.richText = [{
          runs: [{
            text: 'Rollback',
            style: { glow: { color: { kind: 'srgb', value: '000000' }, opacity: 1, size: 2 } },
          }],
        }];
        throw new Error('restore glow');
      }),
    ).toThrow('restore glow');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(slide.shapes[0]).toBe(shape);

    const reopened = await PptxDocument.open(await document.write());
    const edited = reopened.slides[0]!.shapes[0] as ShapeModel;
    const duplicated = reopened.slides[1]!.shapes[0] as ShapeModel;
    expect(edited.richText[0]!.runs.map(({ style }) => style?.glow)).toEqual([
      { color: { kind: 'scheme', value: 'tx1' }, opacity: 0.6, size: 5 },
      undefined,
    ]);
    expect(duplicated.richText[0]!.runs[0]!.style!.glow).toEqual({
      color: { kind: 'srgb', value: 'FF0000' },
      opacity: 0.5,
      size: 8,
    });
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('creates, edits, clears, duplicates, rolls back, and reopens rich text transparency', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addRichText([{
      indent: -12,
      rtl: true,
      spacing: { before: 4, after: 6, line: { kind: 'multiple', factor: 1.25 } },
      tabStops: [{ position: 1.5, alignment: 'decimal' }],
      runs: [
        { text: 'Omitted' },
        { text: 'Opaque', style: { color: { kind: 'srgb', value: 'FF0000' }, transparency: 0 } },
        {
          text: 'Quarter',
          style: {
            fontFamily: 'Aptos',
            fontSize: 18,
            lang: 'fr-CA',
            baseline: 'superscript',
            characterSpacing: 2.5,
            bold: true,
            italic: true,
            color: { kind: 'srgb', value: '00FF00' },
            transparency: 25,
            glow: { color: { kind: 'scheme', value: 'accent3' }, opacity: 0.25, size: 6 },
            highlight: { kind: 'srgb', value: 'FFFF00' },
            outline: { color: { kind: 'scheme', value: 'accent2' }, size: 0.75 },
            underline: { style: 'dbl', color: { kind: 'scheme', value: 'accent1' } },
            strike: 'dblStrike',
          },
        },
        { text: 'Fractional', style: { color: { kind: 'srgb', value: '0000FF' }, transparency: 50.5555 } },
        { text: 'Invisible', style: { color: { kind: 'srgb', value: '112233' }, transparency: 100 } },
        { text: 'Theme', style: { color: { kind: 'scheme', value: 'accent1' }, transparency: 40 } },
        { text: 'Default', style: { transparency: 60 } },
        { text: '', style: { color: { kind: 'scheme', value: 'accent2' }, transparency: 75 } },
      ],
    }]);

    expect(shape.richText[0]!.runs.map(({ style }) => style?.transparency)).toEqual([
      undefined,
      0,
      25,
      50.555,
      100,
      40,
      60,
      75,
    ]);
    let slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(slideXml).toContain('<a:srgbClr val="FF0000"><a:alpha val="100000"/></a:srgbClr>');
    expect(slideXml).toContain('<a:srgbClr val="00FF00"><a:alpha val="75000"/></a:srgbClr>');
    expect(slideXml).toContain('<a:srgbClr val="0000FF"><a:alpha val="49445"/></a:srgbClr>');
    expect(slideXml).toContain('<a:srgbClr val="112233"><a:alpha val="0"/></a:srgbClr>');
    expect(slideXml).toContain('<a:schemeClr val="tx1"><a:alpha val="40000"/></a:schemeClr>');
    expect(slideXml).toContain(
      '<a:glow rad="76200"><a:schemeClr val="accent3"><a:alpha val="25000"/></a:schemeClr></a:glow>',
    );
    expect(slideXml).toContain('<a:ln w="9525"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:ln>');
    expect(slideXml).toContain('<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>');
    expect(slideXml).toContain('<a:uFill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:uFill>');

    const snapshot = shape.richText as unknown as Array<{
      runs: Array<{ style?: { transparency?: number } }>;
    }>;
    snapshot[0]!.runs[1]!.style!.transparency = 50;
    expect(shape.richText[0]!.runs[1]!.style!.transparency).toBe(0);

    document.duplicateSlide(0);
    shape.richText = [{
      runs: [
        { text: 'Ten', style: { transparency: 10 } },
        { text: 'Ninety', style: { color: { kind: 'scheme', value: 'accent2' }, transparency: 90 } },
        { text: 'Zero', style: { color: { kind: 'srgb', value: 'FF0000' }, transparency: 0 } },
        { text: 'Hundred', style: { transparency: 100 } },
        { text: 'Cleared', style: { bold: true } },
      ],
    }];
    expect(shape.richText[0]!.runs.map(({ style }) => style?.transparency)).toEqual([
      10,
      90,
      0,
      100,
      undefined,
    ]);
    slideXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    const clearedIndex = slideXml.indexOf('>Cleared<');
    const clearedStart = slideXml.lastIndexOf('<a:r>', clearedIndex);
    const clearedEnd = slideXml.indexOf('</a:r>', clearedIndex);
    expect(slideXml.slice(clearedStart, clearedEnd)).not.toContain('<a:alpha');

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    const liveState = shape.richText;
    expect(() =>
      document.transaction(() => {
        shape.richText = [{ runs: [{ text: 'Rollback', style: { transparency: 33.3333 } }] }];
        throw new Error('restore transparency');
      }),
    ).toThrow('restore transparency');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(document.slides[0]).toBe(slide);
    expect(slide.shapes[0]).toBe(shape);
    expect(shape.richText).toEqual(liveState);

    const reopened = await PptxDocument.open(await document.write());
    const edited = reopened.slides[0]!.shapes[0] as ShapeModel;
    const duplicated = reopened.slides[1]!.shapes[0] as ShapeModel;
    expect(edited.richText[0]!.runs.map(({ style }) => style?.transparency)).toEqual([
      10,
      90,
      0,
      100,
      undefined,
    ]);
    expect(duplicated.richText[0]!.runs.map(({ style }) => style?.transparency)).toEqual([
      undefined,
      0,
      25,
      50.555,
      100,
      40,
      60,
      75,
    ]);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('rejects malformed rich text transparency without changing package state', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addRichText([{
      runs: [{ text: 'Original', style: { transparency: 25 } }],
    }]);
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];

    for (const transparency of [
      null,
      true,
      false,
      '25',
      {},
      [],
      Symbol('transparency'),
      Number.NaN,
      Number.POSITIVE_INFINITY,
    ]) {
      expect(() => slide.addRichText([{
        runs: [{ text: 'Invalid', style: { transparency: transparency as never } }],
      }])).toThrow(TypeError);
      expect(() => {
        shape.richText = [{
          runs: [{ text: 'Invalid', style: { transparency: transparency as never } }],
        }];
      }).toThrow(TypeError);
    }
    for (const transparency of [-0.001, 100.001]) {
      expect(() => slide.addRichText([{
        runs: [{ text: 'Invalid', style: { transparency } }],
      }])).toThrow(RangeError);
      expect(() => {
        shape.richText = [{ runs: [{ text: 'Invalid', style: { transparency } }] }];
      }).toThrow(RangeError);
    }

    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(slide.shapes).toHaveLength(1);
    expect(slide.shapes[0]).toBe(shape);
    expect(shape.richText[0]!.runs[0]!.style!.transparency).toBe(25);
  });

  it('creates, edits, duplicates, and reopens rich text baselines', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addRichText([{
      runs: [
        { text: 'Super', style: { baseline: 'superscript' } },
        { text: 'Sub', style: { baseline: 'subscript' } },
        { text: 'Normal', style: { baseline: 0 } },
        { text: 'Custom', style: { baseline: 12.3456 } },
        { text: 'Minimum', style: { baseline: -2_147_483.648 } },
        { text: 'Maximum', style: { baseline: 2_147_483.647 } },
        {
          text: 'Combined',
          style: {
            baseline: 30,
            glow: { color: { kind: 'scheme', value: 'accent3' }, opacity: 0.25, size: 6 },
            outline: { color: { kind: 'scheme', value: 'accent2' }, size: 0.75 },
            strike: 'dblStrike',
            underline: { style: 'dbl' },
          },
        },
      ],
    }]);

    expect(shape.richText[0]!.runs.map(({ style }) => style?.baseline)).toEqual([
      'superscript',
      'subscript',
      0,
      12.346,
      -2_147_483.648,
      2_147_483.647,
      'superscript',
    ]);
    const createdXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(createdXml).toContain('<a:rPr lang="en-US" baseline="0" dirty="0">');
    expect(createdXml).toContain('baseline="12346"');
    expect(createdXml).toContain('baseline="-2147483648"');
    expect(createdXml).toContain('baseline="2147483647"');

    document.duplicateSlide(0);
    shape.richText = [{
      runs: [
        { text: 'Changed', style: { baseline: -12.5 } },
        { text: 'Cleared' },
      ],
    }];
    expect(shape.richText[0]!.runs.map(({ style }) => style?.baseline)).toEqual([-12.5, undefined]);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    expect(() =>
      document.transaction(() => {
        shape.richText = [{ runs: [{ text: 'Rollback', style: { baseline: 'subscript' } }] }];
        throw new Error('restore baseline');
      }),
    ).toThrow('restore baseline');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(slide.shapes[0]).toBe(shape);

    const reopened = await PptxDocument.open(await document.write());
    const edited = reopened.slides[0]!.shapes[0] as ShapeModel;
    const duplicated = reopened.slides[1]!.shapes[0] as ShapeModel;
    expect(edited.richText[0]!.runs.map(({ style }) => style?.baseline)).toEqual([-12.5, undefined]);
    expect(duplicated.richText[0]!.runs[0]!.style!.baseline).toBe('superscript');
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('creates, edits, duplicates, and reopens rich text character spacing', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addRichText([{
      runs: [
        { text: 'Expanded', style: { characterSpacing: 2.5 } },
        { text: 'Condensed', style: { characterSpacing: -1.25 } },
        { text: 'Normal', style: { characterSpacing: 0 } },
        { text: 'Quantized', style: { characterSpacing: 0.004 } },
        { text: 'Minimum', style: { characterSpacing: -21_474_836.48 } },
        { text: 'Maximum', style: { characterSpacing: 21_474_836.47 } },
        {
          text: 'Combined',
          style: {
            baseline: 'superscript',
            characterSpacing: 3,
            glow: { color: { kind: 'scheme', value: 'accent3' }, opacity: 0.25, size: 6 },
            outline: { color: { kind: 'scheme', value: 'accent2' }, size: 0.75 },
            strike: 'dblStrike',
            underline: { style: 'dbl' },
          },
        },
      ],
    }]);

    expect(shape.richText[0]!.runs.map(({ style }) => style?.characterSpacing)).toEqual([
      2.5,
      -1.25,
      0,
      0,
      -21_474_836.48,
      21_474_836.47,
      3,
    ]);
    const createdXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(createdXml).toContain('spc="250" kern="0"');
    expect(createdXml.match(/spc="0" kern="0"/g)).toHaveLength(2);
    expect(createdXml).toContain('baseline="30000" spc="300" kern="0"');

    document.duplicateSlide(0);
    shape.richText = [{
      runs: [
        { text: 'Changed', style: { characterSpacing: -2.75 } },
        { text: 'Cleared' },
      ],
    }];
    expect(shape.richText[0]!.runs.map(({ style }) => style?.characterSpacing)).toEqual([-2.75, undefined]);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    expect(() =>
      document.transaction(() => {
        shape.richText = [{ runs: [{ text: 'Rollback', style: { characterSpacing: 4 } }] }];
        throw new Error('restore character spacing');
      }),
    ).toThrow('restore character spacing');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(slide.shapes[0]).toBe(shape);

    const reopened = await PptxDocument.open(await document.write());
    const edited = reopened.slides[0]!.shapes[0] as ShapeModel;
    const duplicated = reopened.slides[1]!.shapes[0] as ShapeModel;
    expect(edited.richText[0]!.runs.map(({ style }) => style?.characterSpacing)).toEqual([-2.75, undefined]);
    expect(duplicated.richText[0]!.runs[0]!.style!.characterSpacing).toBe(2.5);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('creates, edits, duplicates, and reopens plain and rich text languages', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const defaultPlain = slide.addText('Default language', { name: 'Language default' });
    const frenchPlain = slide.addText('Bonjour\n\nEncore', {
      name: 'Language plain',
      lang: 'fr-CA',
    });
    const rich = slide.addRichText([{
      runs: [
        { text: 'Inherited' },
        { text: ' German', style: { lang: 'de-DE' } },
        { text: ' Explicit default', style: { lang: 'en-US' } },
        { text: ' Escaped', style: { bold: true, lang: 'x-private&"quoted' } },
      ],
    }], {
      name: 'Language rich',
      lang: 'fr-CA',
    });
    const escapedPlain = slide.addText('Escaped outer language', {
      name: 'Language escaped outer',
      lang: 'x-private&"quoted',
    });

    expect(defaultPlain.richText[0]!.runs[0]!.style!.lang).toBe('en-US');
    expect(frenchPlain.richText.map((paragraph) => paragraph.runs[0]?.style?.lang)).toEqual([
      'fr-CA',
      undefined,
      'fr-CA',
    ]);
    expect(rich.richText[0]!.runs.map(({ style }) => style?.lang)).toEqual([
      'fr-CA',
      'de-DE',
      'en-US',
      'x-private&"quoted',
    ]);
    expect(escapedPlain.richText[0]!.runs[0]!.style!.lang).toBe('x-private&"quoted');

    const createdXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(createdXml).toMatch(
      /name="Language default"[\s\S]*?<a:rPr lang="en-US" dirty="0">/,
    );
    expect(createdXml).toMatch(
      /name="Language plain"[\s\S]*?<a:rPr lang="fr-CA" altLang="en-US" dirty="0">/,
    );
    expect(createdXml).toContain(
      '<a:rPr lang="x-private&amp;&quot;quoted" altLang="en-US" b="1" dirty="0">',
    );
    expect(createdXml).toContain(
      '<a:endParaRPr lang="x-private&amp;&quot;quoted" dirty="0"/>',
    );
    expect(createdXml.match(/<a:endParaRPr lang="fr-CA" dirty="0"\/>/g)).toHaveLength(4);

    document.duplicateSlide(0);
    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    expect(() =>
      document.transaction(() => {
        rich.richText = [{ runs: [{ text: 'Rollback', style: { lang: 'ko-KR' } }] }];
        throw new Error('restore language');
      }),
    ).toThrow('restore language');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(slide.shapes[2]).toBe(rich);

    rich.richText = [{
      runs: [
        { text: 'Japanese', style: { lang: 'ja-JP' } },
        { text: ' Default' },
      ],
    }];
    expect(rich.richText[0]!.runs.map(({ style }) => style?.lang)).toEqual(['ja-JP', 'en-US']);

    const reopened = await PptxDocument.open(await document.write());
    const edited = reopened.slides[0]!.shapes[2] as ShapeModel;
    const duplicated = reopened.slides[1]!.shapes[2] as ShapeModel;
    expect(edited.richText[0]!.runs.map(({ style }) => style?.lang)).toEqual(['ja-JP', 'en-US']);
    expect(duplicated.richText[0]!.runs.map(({ style }) => style?.lang)).toEqual([
      'fr-CA',
      'de-DE',
      'en-US',
      'x-private&"quoted',
    ]);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('rejects malformed text languages without changing package state', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addRichText([{ runs: [{ text: 'Original', style: { lang: 'en-US' } }] }], {
      lang: 'fr-CA',
    });
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];

    for (const language of [
      null,
      true,
      false,
      0,
      1,
      '',
      {},
      [],
      Symbol('language'),
      'bad\u0000language',
    ]) {
      expect(() => slide.addText('Invalid', { lang: language as never })).toThrow(TypeError);
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid' }] }], {
        lang: language as never,
      })).toThrow(TypeError);
      expect(() => slide.addRichText([{
        runs: [{ text: 'Invalid', style: { lang: language as never } }],
      }])).toThrow(TypeError);
      expect(() => {
        shape.richText = [{ runs: [{ text: 'Invalid', style: { lang: language as never } }] }];
      }).toThrow(TypeError);
    }

    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(shape.richText[0]!.runs[0]!.style!.lang).toBe('en-US');
    expect(slide.shapes[0]).toBe(shape);
  });

  it('creates, edits, duplicates, and reopens text-box margins', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const sourceMargins = { top: -0.5, right: 0, bottom: 0.125, left: 8 };
    const uniform = slide.addText('Uniform', { name: 'Uniform margins', margin: 10 });
    const tuple = slide.addRichText([{ runs: [{ text: 'Tuple' }] }], {
      name: 'Tuple margins',
      margin: [1, 2, 3, 4],
    });
    const named = slide.addText('Named', { name: 'Named margins', margin: sourceMargins });
    const zero = slide.addText('Zero', { name: 'Zero margins', margin: 0 });
    const omitted = slide.addText('Omitted', { name: 'Omitted margins' });
    const boundaries = slide.addText('Boundaries', {
      name: 'Boundary margins',
      margin: {
        top: -2_147_483_648 / 12_700,
        right: 2_147_483_647 / 12_700,
      },
    });
    sourceMargins.left = 99;

    expect(uniform.textMargins).toEqual({ top: 10, right: 10, bottom: 10, left: 10 });
    expect(tuple.textMargins).toEqual({ top: 1, right: 2, bottom: 3, left: 4 });
    expect(named.textMargins).toEqual({ top: -0.5, right: 0, bottom: 1_588 / 12_700, left: 8 });
    expect(zero.textMargins).toEqual({ top: 0, right: 0, bottom: 0, left: 0 });
    expect(omitted.textMargins).toBeUndefined();
    expect(boundaries.textMargins).toEqual({
      top: -2_147_483_648 / 12_700,
      right: 2_147_483_647 / 12_700,
    });

    const createdXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(createdXml).toContain(
      '<a:bodyPr wrap="square" lIns="127000" tIns="127000" rIns="127000" bIns="127000" rtlCol="0" anchor="ctr"/>',
    );
    expect(createdXml).toContain(
      '<a:bodyPr wrap="square" lIns="50800" tIns="12700" rIns="25400" bIns="38100" rtlCol="0" anchor="ctr"/>',
    );
    expect(createdXml).toContain(
      '<a:bodyPr wrap="square" lIns="101600" tIns="-6350" rIns="0" bIns="1588" rtlCol="0" anchor="ctr"/>',
    );
    expect(createdXml).toContain(
      '<a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="ctr"/>',
    );

    document.duplicateSlide(0);
    uniform.textMargins = { top: 4, left: 8 };
    tuple.textMargins = {};
    zero.textMargins = undefined;
    omitted.textMargins = [5, 6, 7, 8];
    expect(uniform.textMargins).toEqual({ top: 4, left: 8 });
    expect(tuple.textMargins).toBeUndefined();
    expect(zero.textMargins).toBeUndefined();
    expect(omitted.textMargins).toEqual({ top: 5, right: 6, bottom: 7, left: 8 });
    expect(slide.shapes[0]).toBe(uniform);
    expect(slide.shapes[1]).toBe(tuple);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    expect(() =>
      document.transaction(() => {
        uniform.textMargins = [9, 8, 7, 6];
        throw new Error('restore text margins');
      }),
    ).toThrow('restore text margins');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(uniform.textMargins).toEqual({ top: 4, left: 8 });
    expect(slide.shapes[0]).toBe(uniform);

    const reopened = await PptxDocument.open(await document.write());
    const editedUniform = reopened.slides[0]!.shapes[0] as ShapeModel;
    const editedTuple = reopened.slides[0]!.shapes[1] as ShapeModel;
    const editedZero = reopened.slides[0]!.shapes[3] as ShapeModel;
    const editedOmitted = reopened.slides[0]!.shapes[4] as ShapeModel;
    const duplicatedUniform = reopened.slides[1]!.shapes[0] as ShapeModel;
    const duplicatedTuple = reopened.slides[1]!.shapes[1] as ShapeModel;
    const duplicatedZero = reopened.slides[1]!.shapes[3] as ShapeModel;
    const duplicatedOmitted = reopened.slides[1]!.shapes[4] as ShapeModel;
    expect(editedUniform.textMargins).toEqual({ top: 4, left: 8 });
    expect(editedTuple.textMargins).toBeUndefined();
    expect(editedZero.textMargins).toBeUndefined();
    expect(editedOmitted.textMargins).toEqual({ top: 5, right: 6, bottom: 7, left: 8 });
    expect(duplicatedUniform.textMargins).toEqual({ top: 10, right: 10, bottom: 10, left: 10 });
    expect(duplicatedTuple.textMargins).toEqual({ top: 1, right: 2, bottom: 3, left: 4 });
    expect(duplicatedZero.textMargins).toEqual({ top: 0, right: 0, bottom: 0, left: 0 });
    expect(duplicatedOmitted.textMargins).toBeUndefined();
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('rejects malformed text-box margins without changing package state', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addText('Original', { margin: { top: 4, left: 8 } });
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    const invalid = [
      null,
      true,
      '1',
      Number.NaN,
      Number.POSITIVE_INFINITY,
      Number.NEGATIVE_INFINITY,
      [-2_147_483_649 / 12_700, 0, 0, 0],
      [2_147_483_648 / 12_700, 0, 0, 0],
      [1, 2, 3],
      [1, 2, 3, 4, 5],
      [1, 2, Number.NaN, 4],
      { top: '1' },
      { right: Number.POSITIVE_INFINITY },
      { inline: 1 },
    ];

    for (const margin of invalid) {
      expect(() => slide.addText('Invalid', { margin: margin as never })).toThrow();
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid' }] }], {
        margin: margin as never,
      })).toThrow();
      expect(() => {
        shape.textMargins = margin as never;
      }).toThrow();
    }
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(shape.text).toBe('Original');
    expect(shape.textMargins).toEqual({ top: 4, left: 8 });
    expect(slide.shapes[0]).toBe(shape);
  });

  it('creates, edits, duplicates, and reopens text-box vertical alignment', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const omitted = slide.addText('Omitted', { name: 'Omitted vertical alignment' });
    const top = slide.addText('Top', {
      name: 'Top vertical alignment',
      valign: 'top',
      margin: 4,
    });
    const middle = slide.addRichText([{
      runs: [{ text: 'Middle rich text', style: { bold: true } }],
      align: 'center',
    }], {
      name: 'Middle vertical alignment',
      valign: 'middle',
      margin: [1, 2, 3, 4],
    });
    const bottom = slide.addText('Bottom', {
      name: 'Bottom vertical alignment',
      valign: 'bottom',
    });

    expect([omitted, top, middle, bottom].map(({ verticalAlignment }) => verticalAlignment)).toEqual([
      'middle',
      'top',
      'middle',
      'bottom',
    ]);
    expect(middle.textMargins).toEqual({ top: 1, right: 2, bottom: 3, left: 4 });
    const createdXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(createdXml).toContain(
      '<a:bodyPr wrap="square" lIns="50800" tIns="50800" rIns="50800" bIns="50800" rtlCol="0" anchor="t"/>',
    );
    expect(createdXml).toContain('<a:bodyPr wrap="square" rtlCol="0" anchor="b"/>');

    document.duplicateSlide(0);
    top.verticalAlignment = 'bottom';
    top.text = 'Plain replacement';
    top.richText = [{ runs: [{ text: 'Rich replacement', style: { italic: true } }] }];
    top.textMargins = { top: 6, left: 8 };
    top.setTransform({ x: inches(2) });
    middle.verticalAlignment = undefined;
    expect(middle.verticalAlignment).toBeUndefined();
    middle.verticalAlignment = 'top';
    expect(middle.verticalAlignment).toBe('top');
    expect(new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes))
      .toMatch(/name="Middle vertical alignment"[\s\S]*?<a:bodyPr[^>]*anchor="t"\/>/);
    middle.verticalAlignment = undefined;
    bottom.verticalAlignment = 'top';
    expect(top.verticalAlignment).toBe('bottom');
    expect(middle.verticalAlignment).toBeUndefined();
    expect(bottom.verticalAlignment).toBe('top');
    expect(slide.shapes[1]).toBe(top);
    expect(slide.shapes[2]).toBe(middle);

    const editedXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(editedXml).toContain('lIns="101600" tIns="76200" rtlCol="0" anchor="b"/>');
    expect(editedXml).toMatch(/name="Middle vertical alignment"[\s\S]*?<a:bodyPr[^>]*rtlCol="0"\/>/);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    expect(() =>
      document.transaction(() => {
        bottom.verticalAlignment = 'bottom';
        throw new Error('restore vertical alignment');
      }),
    ).toThrow('restore vertical alignment');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(bottom.verticalAlignment).toBe('top');
    expect(slide.shapes[3]).toBe(bottom);

    const reopened = await PptxDocument.open(await document.write());
    const edited = reopened.slides[0]!.shapes as readonly ShapeModel[];
    const duplicated = reopened.slides[1]!.shapes as readonly ShapeModel[];
    expect(edited.map(({ verticalAlignment }) => verticalAlignment)).toEqual([
      'middle',
      'bottom',
      undefined,
      'top',
    ]);
    expect(duplicated.map(({ verticalAlignment }) => verticalAlignment)).toEqual([
      'middle',
      'top',
      'middle',
      'bottom',
    ]);
    expect(edited[1]!.text).toBe('Rich replacement');
    expect(edited[1]!.textMargins).toEqual({ top: 6, left: 8 });
    expect(edited[2]!.textMargins).toEqual({ top: 1, right: 2, bottom: 3, left: 4 });
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('rejects malformed text-box vertical alignment without changing package state', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addText('Original', { valign: 'top' });
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];

    for (const valign of [
      null,
      true,
      1,
      '',
      'top ',
      'Top',
      'center',
      't',
      'ctr',
      'b',
      'just',
      'dist',
      {},
      [],
      Symbol('top'),
    ]) {
      expect(() => slide.addText('Invalid', { valign: valign as never })).toThrow(TypeError);
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid' }] }], {
        valign: valign as never,
      })).toThrow(TypeError);
      expect(() => {
        shape.verticalAlignment = valign as never;
      }).toThrow(TypeError);
    }
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(shape.text).toBe('Original');
    expect(shape.verticalAlignment).toBe('top');
    expect(slide.shapes[0]).toBe(shape);
  });

  it('creates, edits, duplicates, and reopens text-box wrapping', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const omitted = slide.addText('Omitted', { name: 'Omitted text wrapping' });
    const wrapped = slide.addRichText([{
      runs: [{ text: 'Wrapped rich text', style: { bold: true } }],
      align: 'center',
    }], {
      name: 'Enabled text wrapping',
      wrap: true,
      margin: 8,
      valign: 'bottom',
    });
    const unwrapped = slide.addText('Unwrapped', {
      name: 'Disabled text wrapping',
      wrap: false,
      valign: 'top',
    });

    expect([omitted, wrapped, unwrapped].map(({ textWrap }) => textWrap)).toEqual([
      true,
      true,
      false,
    ]);
    expect(wrapped.textMargins).toEqual({ top: 8, right: 8, bottom: 8, left: 8 });
    expect(wrapped.verticalAlignment).toBe('bottom');
    const createdXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(createdXml).toContain(
      '<a:bodyPr wrap="square" lIns="101600" tIns="101600" rIns="101600" bIns="101600" rtlCol="0" anchor="b"/>',
    );
    expect(createdXml).toContain('<a:bodyPr wrap="none" rtlCol="0" anchor="t"/>');

    document.duplicateSlide(0);
    omitted.textWrap = false;
    wrapped.textWrap = false;
    wrapped.text = 'Plain replacement';
    wrapped.richText = [{ runs: [{ text: 'Rich replacement', style: { italic: true } }] }];
    wrapped.textMargins = { top: 6, left: 8 };
    wrapped.verticalAlignment = 'middle';
    wrapped.setTransform({ x: inches(2) });
    unwrapped.textWrap = undefined;
    expect(unwrapped.textWrap).toBeUndefined();
    unwrapped.textWrap = true;
    expect(unwrapped.textWrap).toBe(true);
    expect(new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes))
      .toMatch(/name="Disabled text wrapping"[\s\S]*?<a:bodyPr[^>]*wrap="square"[^>]*\/>/);
    unwrapped.textWrap = undefined;

    expect(omitted.textWrap).toBe(false);
    expect(wrapped.textWrap).toBe(false);
    expect(unwrapped.textWrap).toBeUndefined();
    expect(slide.shapes[0]).toBe(omitted);
    expect(slide.shapes[1]).toBe(wrapped);
    expect(slide.shapes[2]).toBe(unwrapped);
    const editedXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(editedXml).toMatch(/name="Enabled text wrapping"[\s\S]*?<a:bodyPr[^>]*wrap="none"[^>]*anchor="ctr"\/>/);
    expect(editedXml).toMatch(/name="Disabled text wrapping"[\s\S]*?<a:bodyPr(?![^>]*\swrap=)[^>]*anchor="t"\/>/);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    expect(() =>
      document.transaction(() => {
        wrapped.textWrap = true;
        throw new Error('restore text wrapping');
      }),
    ).toThrow('restore text wrapping');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(wrapped.textWrap).toBe(false);
    expect(slide.shapes[1]).toBe(wrapped);

    const reopened = await PptxDocument.open(await document.write());
    const edited = reopened.slides[0]!.shapes as readonly ShapeModel[];
    const duplicated = reopened.slides[1]!.shapes as readonly ShapeModel[];
    expect(edited.map(({ textWrap }) => textWrap)).toEqual([false, false, undefined]);
    expect(duplicated.map(({ textWrap }) => textWrap)).toEqual([true, true, false]);
    expect(edited[1]!.text).toBe('Rich replacement');
    expect(edited[1]!.textMargins).toEqual({ top: 6, left: 8 });
    expect(edited[1]!.verticalAlignment).toBe('middle');
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('rejects malformed text-box wrapping without changing package state', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addText('Original', {
      wrap: false,
      margin: { top: 4, left: 8 },
      valign: 'top',
    });
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];

    for (const wrap of [
      null,
      0,
      1,
      '',
      'true',
      'false',
      'square',
      'none',
      {},
      [],
      Symbol('wrap'),
    ]) {
      expect(() => slide.addText('Invalid', { wrap: wrap as never })).toThrow(TypeError);
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid' }] }], {
        wrap: wrap as never,
      })).toThrow(TypeError);
      expect(() => {
        shape.textWrap = wrap as never;
      }).toThrow(TypeError);
    }
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(shape.text).toBe('Original');
    expect(shape.textMargins).toEqual({ top: 4, left: 8 });
    expect(shape.verticalAlignment).toBe('top');
    expect(shape.textWrap).toBe(false);
    expect(slide.shapes[0]).toBe(shape);
  });

  it('creates, edits, duplicates, and reopens every text-box text direction', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const directions = [
      'eaVert',
      'horz',
      'mongolianVert',
      'vert',
      'vert270',
      'wordArtVert',
      'wordArtVertRtl',
    ] as const;
    const omitted = slide.addText('Omitted', { name: 'Direction omitted' });
    const directed = directions.map((direction, index) => index % 2 === 0
      ? slide.addText(direction, {
        name: `Direction ${direction}`,
        vert: direction,
        ...(index === 0 ? { margin: 8, valign: 'bottom' as const, wrap: false } : {}),
      })
      : slide.addRichText([{
        align: 'center',
        runs: [{ text: direction, style: { bold: true } }],
      }], {
        name: `Direction ${direction}`,
        vert: direction,
      }));

    expect(omitted.textDirection).toBeUndefined();
    expect(directed.map(({ textDirection }) => textDirection)).toEqual(directions);
    expect(directed[0]!.textMargins).toEqual({ top: 8, right: 8, bottom: 8, left: 8 });
    expect(directed[0]!.verticalAlignment).toBe('bottom');
    expect(directed[0]!.textWrap).toBe(false);
    const createdXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(createdXml).toMatch(
      /name="Direction omitted"[\s\S]*?<a:bodyPr(?![^>]*\svert=)[^>]*\/>/,
    );
    for (const direction of directions) {
      expect(createdXml).toMatch(
        new RegExp(`name="Direction ${direction}"[\\s\\S]*?<a:bodyPr[^>]* vert="${direction}"\\/>`),
      );
    }

    document.duplicateSlide(0);
    directed[0]!.textDirection = 'vert270';
    directed[0]!.text = 'Plain replacement';
    directed[0]!.richText = [{ runs: [{ text: 'Rich replacement', style: { italic: true } }] }];
    directed[0]!.textMargins = { top: 6, left: 8 };
    directed[0]!.verticalAlignment = 'middle';
    directed[0]!.textWrap = true;
    directed[0]!.setTransform({ x: inches(2) });
    directed[1]!.textDirection = undefined;
    expect(directed[1]!.textDirection).toBeUndefined();
    directed[1]!.textDirection = 'wordArtVert';
    expect(directed[1]!.textDirection).toBe('wordArtVert');
    directed[1]!.textDirection = undefined;
    omitted.textDirection = 'horz';
    expect(omitted.textDirection).toBe('horz');
    omitted.textDirection = undefined;
    expect(omitted.textDirection).toBeUndefined();
    expect(slide.shapes[0]).toBe(omitted);
    expect(slide.shapes[1]).toBe(directed[0]);
    expect(slide.shapes[2]).toBe(directed[1]);

    const editedXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(editedXml).toMatch(/name="Direction eaVert"[\s\S]*?<a:bodyPr[^>]* vert="vert270"\/>/);
    expect(editedXml).toMatch(
      /name="Direction horz"[\s\S]*?<a:bodyPr(?![^>]*\svert=)[^>]*\/>/,
    );

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    expect(() =>
      document.transaction(() => {
        directed[2]!.textDirection = 'wordArtVertRtl';
        throw new Error('restore text direction');
      }),
    ).toThrow('restore text direction');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(directed[2]!.textDirection).toBe('mongolianVert');
    expect(slide.shapes[3]).toBe(directed[2]);

    const reopened = await PptxDocument.open(await document.write());
    const edited = reopened.slides[0]!.shapes as readonly ShapeModel[];
    const duplicated = reopened.slides[1]!.shapes as readonly ShapeModel[];
    expect(edited.map(({ textDirection }) => textDirection)).toEqual([
      undefined,
      'vert270',
      undefined,
      'mongolianVert',
      'vert',
      'vert270',
      'wordArtVert',
      'wordArtVertRtl',
    ]);
    expect(duplicated.map(({ textDirection }) => textDirection)).toEqual([
      undefined,
      ...directions,
    ]);
    expect(edited[1]!.text).toBe('Rich replacement');
    expect(edited[1]!.textMargins).toEqual({ top: 6, left: 8 });
    expect(edited[1]!.verticalAlignment).toBe('middle');
    expect(edited[1]!.textWrap).toBe(true);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('rejects malformed text-box text direction without changing package state', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addText('Original', {
      vert: 'vert',
      margin: { top: 4, left: 8 },
      valign: 'top',
      wrap: false,
    });
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];

    for (const direction of [
      null,
      true,
      1,
      '',
      'Vert',
      ' vert ',
      'vertical',
      'unknown',
      {},
      [],
      Symbol('vert'),
    ]) {
      expect(() => slide.addText('Invalid', { vert: direction as never })).toThrow(TypeError);
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid' }] }], {
        vert: direction as never,
      })).toThrow(TypeError);
      expect(() => {
        shape.textDirection = direction as never;
      }).toThrow(TypeError);
    }
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(shape.text).toBe('Original');
    expect(shape.textMargins).toEqual({ top: 4, left: 8 });
    expect(shape.verticalAlignment).toBe('top');
    expect(shape.textWrap).toBe(false);
    expect(shape.textDirection).toBe('vert');
    expect(slide.shapes[0]).toBe(shape);
  });

  it('creates, edits, duplicates, and reopens text-box fit modes', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const omitted = slide.addText('Omitted', { name: 'Fit omitted' });
    const none = slide.addText('None', { name: 'Fit none', fit: 'none' });
    const shrink = slide.addRichText([{
      align: 'center',
      runs: [{ text: 'Shrink rich text', style: { bold: true } }],
    }], {
      name: 'Fit shrink',
      fit: 'shrink',
      margin: 8,
      valign: 'bottom',
      vert: 'vert',
      wrap: false,
    });
    const resize = slide.addText('Resize', { name: 'Fit resize', fit: 'resize' });

    expect([omitted, none, shrink, resize].map(({ textFit }) => textFit)).toEqual([
      undefined,
      undefined,
      'shrink',
      'resize',
    ]);
    expect(shrink.textMargins).toEqual({ top: 8, right: 8, bottom: 8, left: 8 });
    expect(shrink.verticalAlignment).toBe('bottom');
    expect(shrink.textDirection).toBe('vert');
    expect(shrink.textWrap).toBe(false);
    const createdXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(createdXml).toMatch(/name="Fit omitted"[\s\S]*?<a:bodyPr[^>]*\/>/);
    expect(createdXml).toMatch(/name="Fit none"[\s\S]*?<a:bodyPr[^>]*\/>/);
    expect(createdXml).toMatch(
      /name="Fit shrink"[\s\S]*?<a:bodyPr[^>]*><a:normAutofit\/><\/a:bodyPr>/,
    );
    expect(createdXml).toMatch(
      /name="Fit resize"[\s\S]*?<a:bodyPr[^>]*><a:spAutoFit\/><\/a:bodyPr>/,
    );

    document.duplicateSlide(0);
    shrink.textFit = 'resize';
    shrink.text = 'Plain replacement';
    shrink.richText = [{ runs: [{ text: 'Rich replacement', style: { italic: true } }] }];
    shrink.textMargins = { top: 6, left: 8 };
    shrink.verticalAlignment = 'middle';
    shrink.textWrap = true;
    shrink.textDirection = 'vert270';
    shrink.setTransform({ x: inches(2) });
    resize.textFit = undefined;
    expect(resize.textFit).toBeUndefined();
    resize.textFit = 'shrink';
    expect(resize.textFit).toBe('shrink');
    resize.textFit = 'none';
    expect(resize.textFit).toBeUndefined();
    omitted.textFit = 'shrink';
    expect(omitted.textFit).toBe('shrink');
    omitted.textFit = 'none';
    expect(omitted.textFit).toBeUndefined();
    expect(slide.shapes[0]).toBe(omitted);
    expect(slide.shapes[2]).toBe(shrink);
    expect(slide.shapes[3]).toBe(resize);

    const editedXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(editedXml).toMatch(
      /name="Fit shrink"[\s\S]*?<a:bodyPr[^>]*><a:spAutoFit\/><\/a:bodyPr>/,
    );
    expect(editedXml).toMatch(
      /name="Fit resize"[\s\S]*?<a:bodyPr[^>]*><\/a:bodyPr>/,
    );

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    expect(() =>
      document.transaction(() => {
        shrink.textFit = 'shrink';
        throw new Error('restore text fit');
      }),
    ).toThrow('restore text fit');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(shrink.textFit).toBe('resize');
    expect(slide.shapes[2]).toBe(shrink);

    const reopened = await PptxDocument.open(await document.write());
    const edited = reopened.slides[0]!.shapes as readonly ShapeModel[];
    const duplicated = reopened.slides[1]!.shapes as readonly ShapeModel[];
    expect(edited.map(({ textFit }) => textFit)).toEqual([
      undefined,
      undefined,
      'resize',
      undefined,
    ]);
    expect(duplicated.map(({ textFit }) => textFit)).toEqual([
      undefined,
      undefined,
      'shrink',
      'resize',
    ]);
    expect(edited[2]!.text).toBe('Rich replacement');
    expect(edited[2]!.textMargins).toEqual({ top: 6, left: 8 });
    expect(edited[2]!.verticalAlignment).toBe('middle');
    expect(edited[2]!.textWrap).toBe(true);
    expect(edited[2]!.textDirection).toBe('vert270');
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('rejects malformed text-box fit without changing package state', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addText('Original', {
      fit: 'shrink',
      margin: { top: 4, left: 8 },
      valign: 'top',
      vert: 'vert',
      wrap: false,
    });
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];

    for (const fit of [
      null,
      true,
      false,
      0,
      1,
      '',
      'None',
      ' shrink ',
      'SHRINK',
      'autofit',
      'normal',
      {},
      [],
      Symbol('fit'),
    ]) {
      expect(() => slide.addText('Invalid', { fit: fit as never })).toThrow(TypeError);
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid' }] }], {
        fit: fit as never,
      })).toThrow(TypeError);
      expect(() => {
        shape.textFit = fit as never;
      }).toThrow(TypeError);
    }
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
    expect(shape.text).toBe('Original');
    expect(shape.textMargins).toEqual({ top: 4, left: 8 });
    expect(shape.verticalAlignment).toBe('top');
    expect(shape.textWrap).toBe(false);
    expect(shape.textDirection).toBe('vert');
    expect(shape.textFit).toBe('shrink');
    expect(slide.shapes[0]).toBe(shape);
  });

  it('creates, edits, duplicates, and reopens rich text highlight colors', async () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const sourceHighlight = { kind: 'srgb' as const, value: '#ffff00' };
    const shape = slide.addRichText([{
      runs: [
        { text: 'Yellow', style: { highlight: sourceHighlight } },
        { text: 'Theme', style: { highlight: { kind: 'scheme', value: 'accent2' } } },
        {
          text: 'Combined',
          style: {
            highlight: { kind: 'srgb', value: '00ff00' },
            strike: 'dblStrike',
            underline: { style: 'dbl', color: { kind: 'scheme', value: 'accent1' } },
          },
        },
      ],
    }]);
    sourceHighlight.value = '000000';

    expect(shape.richText[0]!.runs.map(({ style }) => style?.highlight)).toEqual([
      { kind: 'srgb', value: 'FFFF00' },
      { kind: 'scheme', value: 'accent2' },
      { kind: 'srgb', value: '00FF00' },
    ]);
    const createdXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(createdXml).toContain('<a:highlight><a:schemeClr val="accent2"/></a:highlight>');
    const highlightIndex = createdXml.indexOf('<a:highlight>');
    expect(highlightIndex).toBeGreaterThan(createdXml.lastIndexOf('</a:solidFill>', highlightIndex));
    expect(createdXml).toContain(
      '<a:highlight><a:srgbClr val="00FF00"/></a:highlight><a:uFill><a:solidFill>',
    );

    const snapshot = shape.richText as unknown as Array<{
      runs: Array<{ style?: { highlight?: { value: string } } }>;
    }>;
    snapshot[0]!.runs[0]!.style!.highlight!.value = 'FF0000';
    expect(shape.richText[0]!.runs[0]!.style!.highlight).toEqual({ kind: 'srgb', value: 'FFFF00' });

    document.duplicateSlide(0);
    shape.richText = [{
      runs: [
        { text: 'Changed', style: { highlight: { kind: 'scheme', value: 'tx2' } } },
        { text: 'Cleared' },
      ],
    }];
    expect(shape.richText[0]!.runs.map(({ style }) => style?.highlight)).toEqual([
      { kind: 'scheme', value: 'tx2' },
      undefined,
    ]);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    expect(() =>
      document.transaction(() => {
        shape.richText = [{
          runs: [{ text: 'Rollback', style: { highlight: { kind: 'srgb', value: 'FF0000' } } }],
        }];
        throw new Error('restore highlight');
      }),
    ).toThrow('restore highlight');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(slide.shapes[0]).toBe(shape);

    const reopened = await PptxDocument.open(await document.write());
    const edited = reopened.slides[0]!.shapes[0] as ShapeModel;
    const duplicated = reopened.slides[1]!.shapes[0] as ShapeModel;
    expect(edited.richText[0]!.runs.map(({ style }) => style?.highlight)).toEqual([
      { kind: 'scheme', value: 'tx2' },
      undefined,
    ]);
    expect(duplicated.richText[0]!.runs[0]!.style!.highlight).toEqual({
      kind: 'srgb',
      value: 'FFFF00',
    });
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('creates, edits, duplicates, and reopens rich text strike styles', async () => {
    const strikeStyles: readonly RichTextStrikeStyle[] = ['sngStrike', 'dblStrike'];
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addRichText([{
      runs: [
        { text: 'True', style: { strike: true } },
        { text: 'False', style: { strike: false } },
        ...strikeStyles.map((strike) => ({ text: strike, style: { strike } })),
        {
          text: 'Combined',
          style: { bold: true, strike: 'dblStrike', underline: { style: 'wavy' } },
        },
      ],
    }]);

    expect(shape.richText[0]!.runs.map(({ style }) => style?.strike)).toEqual([
      'sngStrike',
      false,
      'sngStrike',
      'dblStrike',
      'dblStrike',
    ]);
    const createdXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(createdXml).toContain('strike="noStrike"');
    expect(createdXml).toContain('b="1" strike="dblStrike" u="wavy"');

    document.duplicateSlide(0);
    shape.richText = [{
      runs: [
        { text: 'Disabled', style: { strike: false } },
        { text: 'Cleared' },
      ],
    }];
    expect(shape.richText[0]!.runs.map(({ style }) => style?.strike)).toEqual([false, undefined]);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    expect(() =>
      document.transaction(() => {
        shape.richText = [{ runs: [{ text: 'Rollback', style: { strike: 'dblStrike' } }] }];
        throw new Error('restore strike');
      }),
    ).toThrow('restore strike');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(slide.shapes[0]).toBe(shape);

    const reopened = await PptxDocument.open(await document.write());
    const edited = reopened.slides[0]!.shapes[0] as ShapeModel;
    const duplicated = reopened.slides[1]!.shapes[0] as ShapeModel;
    expect(edited.richText[0]!.runs.map(({ style }) => style?.strike)).toEqual([false, undefined]);
    expect(duplicated.richText[0]!.runs.map(({ style }) => style?.strike)).toEqual([
      'sngStrike',
      false,
      'sngStrike',
      'dblStrike',
      'dblStrike',
    ]);
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('creates edits duplicates and reopens rich text effects in all six formats', async () => {
    const initialParagraphs: readonly RichTextParagraph[] = [{
      runs: [
        {
          text: 'Primary',
          style: {
            baseline: 'superscript',
            characterSpacing: 2.5,
            color: { kind: 'srgb', value: '112233' },
            glow: {
              color: { kind: 'scheme', value: 'accent1' },
              opacity: 0.5,
              size: 8,
            },
            outline: { color: { kind: 'srgb', value: 'FF0000' }, size: 1.5 },
            strike: 'sngStrike',
            transparency: 25,
          },
        },
        {
          text: ' Secondary',
          style: {
            baseline: 'subscript',
            characterSpacing: -1.25,
            glow: {
              color: { kind: 'srgb', value: '00FF00' },
              opacity: 0,
              size: 0,
            },
            outline: { color: { kind: 'scheme', value: 'accent2' }, size: 0 },
            strike: 'dblStrike',
            transparency: 50.555,
          },
        },
        {
          text: ' Explicit zero',
          style: {
            baseline: 0,
            characterSpacing: 0,
            strike: false,
            transparency: 0,
          },
        },
      ],
    }];
    const editedParagraphs: readonly RichTextParagraph[] = [{
      runs: [
        {
          text: 'Edited',
          style: {
            baseline: 12.345,
            characterSpacing: 1,
            glow: {
              color: { kind: 'srgb', value: '0000FF' },
              opacity: 0.75,
              size: 3,
            },
            outline: { color: { kind: 'scheme', value: 'tx1' }, size: 2 },
            strike: false,
            transparency: 100,
          },
        },
        { text: ' Cleared' },
      ],
    }];
    const snapshot = (shape: ShapeModel) => shape.richText[0]!.runs.map(({ style }) => ({
      baseline: style?.baseline,
      characterSpacing: style?.characterSpacing,
      glow: style?.glow,
      outline: style?.outline,
      strike: style?.strike,
      transparency: style?.transparency,
    }));

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const document = PptxDocument.create({ format });
      const slide = document.addSlide();
      const shape = slide.addRichText(initialParagraphs, {
        name: 'rich-text-effects-family',
      });
      const initialSnapshot = snapshot(shape);
      expect(initialSnapshot).toEqual([
        {
          baseline: 'superscript',
          characterSpacing: 2.5,
          glow: {
            color: { kind: 'scheme', value: 'accent1' },
            opacity: 0.5,
            size: 8,
          },
          outline: { color: { kind: 'srgb', value: 'FF0000' }, size: 1.5 },
          strike: 'sngStrike',
          transparency: 25,
        },
        {
          baseline: 'subscript',
          characterSpacing: -1.25,
          glow: {
            color: { kind: 'srgb', value: '00FF00' },
            opacity: 0,
            size: 0,
          },
          outline: { color: { kind: 'scheme', value: 'accent2' }, size: 0 },
          strike: 'dblStrike',
          transparency: 50.555,
        },
        {
          baseline: 0,
          characterSpacing: 0,
          glow: undefined,
          outline: undefined,
          strike: false,
          transparency: 0,
        },
      ]);
      const createdXml = new TextDecoder().decode(
        document.opcPackage.requirePart(slide.partUri).bytes,
      );
      expect(createdXml).toMatch(
        /<a:rPr\b(?=[^>]*baseline="30000")(?=[^>]*spc="250")(?=[^>]*strike="sngStrike")[^>]*>/u,
      );
      expect(createdXml).toContain(
        '<a:glow rad="101600"><a:schemeClr val="accent1"><a:alpha val="50000"/>',
      );
      expect(createdXml).toContain(
        '<a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>',
      );
      expect(createdXml).toMatch(
        /<a:rPr\b(?=[^>]*baseline="-40000")(?=[^>]*spc="-125")(?=[^>]*strike="dblStrike")[^>]*>/u,
      );
      expect(createdXml).toMatch(
        /<a:rPr\b(?=[^>]*baseline="0")(?=[^>]*spc="0")(?=[^>]*strike="noStrike")[^>]*>/u,
      );

      document.duplicateSlide(0);
      shape.richText = editedParagraphs;
      const editedSnapshot = snapshot(shape);
      expect(editedSnapshot).toEqual([
        {
          baseline: 12.345,
          characterSpacing: 1,
          glow: {
            color: { kind: 'srgb', value: '0000FF' },
            opacity: 0.75,
            size: 3,
          },
          outline: { color: { kind: 'scheme', value: 'tx1' }, size: 2 },
          strike: false,
          transparency: 100,
        },
        {
          baseline: undefined,
          characterSpacing: undefined,
          glow: undefined,
          outline: undefined,
          strike: undefined,
          transparency: undefined,
        },
      ]);
      const editedXml = new TextDecoder().decode(
        document.opcPackage.requirePart(slide.partUri).bytes,
      );
      const clearedRunXml = editedXml.match(
        /<a:r>(?:(?!<\/a:r>)[\s\S])*?<a:t xml:space="preserve"> Cleared<\/a:t><\/a:r>/u,
      )?.[0];
      expect(clearedRunXml).toBeDefined();
      expect(clearedRunXml).not.toMatch(
        /\bbaseline=|\bspc=|\bstrike=|<a:ln\b|<a:effectLst\b|<a:alpha\b/u,
      );
      expect(document.diagnostics.filter(
        ({ severity }) => severity === 'error' || severity === 'warning',
      )).toEqual([]);
      expect(validatePackage(document.opcPackage).filter(
        ({ severity }) => severity === 'error' || severity === 'warning',
      )).toEqual([]);

      const reopened = await PptxDocument.open(await document.write());
      const reopenedSource = reopened.slides[0]!.shapes.find(
        ({ name }) => name === 'rich-text-effects-family',
      );
      const reopenedDuplicate = reopened.slides[1]!.shapes.find(
        ({ name }) => name === 'rich-text-effects-family',
      );
      expect(reopenedSource).toBeInstanceOf(ShapeModel);
      expect(reopenedDuplicate).toBeInstanceOf(ShapeModel);
      expect(snapshot(reopenedSource as ShapeModel)).toEqual(editedSnapshot);
      expect(snapshot(reopenedDuplicate as ShapeModel)).toEqual(initialSnapshot);
      expect(reopened.diagnostics.filter(
        ({ severity }) => severity === 'error' || severity === 'warning',
      )).toEqual([]);
      expect(validatePackage(reopened.opcPackage).filter(
        ({ severity }) => severity === 'error' || severity === 'warning',
      )).toEqual([]);
    }
  });

  it('creates, edits, duplicates, and reopens rich text underline styles', async () => {
    const underlineStyles: readonly RichTextUnderlineStyle[] = [
      'words',
      'sng',
      'dbl',
      'heavy',
      'dotted',
      'dottedHeavy',
      'dash',
      'dashHeavy',
      'dashLong',
      'dashLongHeavy',
      'dotDash',
      'dotDashHeavy',
      'dotDotDash',
      'dotDotDashHeavy',
      'wavy',
      'wavyHeavy',
      'wavyDbl',
    ];
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const sourceColor = { kind: 'srgb' as const, value: '#ff0000' };
    const shape = slide.addRichText([{
      runs: [
        { text: 'True', style: { underline: true } },
        { text: 'False', style: { underline: false } },
        { text: 'Color', style: { underline: { color: sourceColor } } },
        {
          text: 'Scheme',
          style: { underline: { style: 'dbl', color: { kind: 'scheme', value: 'accent2' } } },
        },
        ...underlineStyles.map((style) => ({ text: style, style: { underline: { style } } })),
      ],
    }]);
    sourceColor.value = '000000';

    const underlines = shape.richText[0]!.runs.map(({ style }) => style?.underline);
    expect(underlines.slice(0, 4)).toEqual([
      { style: 'sng' },
      false,
      { style: 'sng', color: { kind: 'srgb', value: 'FF0000' } },
      { style: 'dbl', color: { kind: 'scheme', value: 'accent2' } },
    ]);
    expect(underlines.slice(4).map((underline) =>
      typeof underline === 'object' ? underline.style : underline)).toEqual(underlineStyles);

    const createdXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(createdXml).toContain('u="none"');
    expect(createdXml).toContain('u="dotDashHeavy"');
    expect(createdXml).toContain(
      '<a:uFill><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:uFill>',
    );
    const underlineFillIndex = createdXml.indexOf('<a:uFill>');
    expect(underlineFillIndex).toBeGreaterThan(createdXml.lastIndexOf('<a:solidFill>', underlineFillIndex));
    expect(underlineFillIndex).toBeLessThan(createdXml.indexOf('<a:latin', underlineFillIndex));

    const snapshot = shape.richText as unknown as Array<{
      runs: Array<{ style?: { underline?: { style?: string; color?: { value: string } } } }>;
    }>;
    snapshot[0]!.runs[2]!.style!.underline!.style = 'wavy';
    snapshot[0]!.runs[2]!.style!.underline!.color!.value = '00FF00';
    expect(shape.richText[0]!.runs[2]!.style!.underline).toEqual({
      style: 'sng',
      color: { kind: 'srgb', value: 'FF0000' },
    });

    document.duplicateSlide(0);
    shape.richText = [{
      runs: [
        { text: 'Disabled', style: { underline: false } },
        { text: 'Cleared' },
      ],
    }];
    expect(shape.richText[0]!.runs.map(({ style }) => style?.underline)).toEqual([false, undefined]);

    const beforeRollback = document.opcPackage.requirePart(slide.partUri).bytes;
    expect(() =>
      document.transaction(() => {
        shape.richText = [{ runs: [{ text: 'Rollback', style: { underline: { style: 'wavyDbl' } } }] }];
        throw new Error('restore underline');
      }),
    ).toThrow('restore underline');
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(slide.shapes[0]).toBe(shape);

    const reopened = await PptxDocument.open(await document.write());
    const edited = reopened.slides[0]!.shapes[0] as ShapeModel;
    const duplicated = reopened.slides[1]!.shapes[0] as ShapeModel;
    expect(edited.richText[0]!.runs.map(({ style }) => style?.underline)).toEqual([false, undefined]);
    expect(duplicated.richText[0]!.runs[2]!.style!.underline).toEqual({
      style: 'sng',
      color: { kind: 'srgb', value: 'FF0000' },
    });
    expect(duplicated.richText[0]!.runs.at(-1)!.style!.underline).toEqual({ style: 'wavyDbl' });
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
  });

  it('creates edits clears and reopens underline owners in all six formats', async () => {
    const commonStyles: readonly RichTextUnderlineStyle[] = [
      'dash',
      'dashHeavy',
      'dashLong',
      'dashLongHeavy',
      'dbl',
      'dotDash',
      'dotDotDash',
      'dotDotDashHeavy',
      'dotted',
      'dottedHeavy',
      'heavy',
      'sng',
      'wavy',
      'wavyDbl',
      'wavyHeavy',
    ];
    const expectedTokens = [
      ...commonStyles,
      'sng',
      'none',
      'dbl',
      'sng',
      'words',
      'dotDashHeavy',
    ];
    const underlineRuns = () => [
      ...commonStyles.map((style) => ({
        text: style,
        style: { underline: { style } },
      })),
      { text: 'true', style: { underline: true as const } },
      { text: 'none', style: { underline: false as const } },
      {
        text: 'srgb',
        style: {
          underline: {
            style: 'dbl' as const,
            color: { kind: 'srgb' as const, value: 'FF0000' },
          },
        },
      },
      {
        text: 'scheme',
        style: {
          underline: { color: { kind: 'scheme' as const, value: 'accent2' as const } },
        },
      },
      { text: 'words', style: { underline: { style: 'words' as const } } },
      {
        text: 'dotDashHeavy',
        style: { underline: { style: 'dotDashHeavy' as const } },
      },
    ];
    const underlineSnapshot = (
      paragraphs: readonly RichTextParagraph[],
    ) => paragraphs[0]!.runs.map(({ style }) => style?.underline);
    const underlineTokens = (xml: string) => [...xml.matchAll(
      /<a:rPr\b[^>]*\bu="([^"]+)"/gu,
    )].map((match) => match[1]);
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

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({ format });
      const layout = created.layouts[0]!;
      const placeholder = layout.addPlaceholder('Placeholder underline owner', {
        name: 'underline-placeholder-owner',
        type: 'body',
        index: 302,
      });
      placeholder.richText = [{ runs: underlineRuns() }];
      const slide = created.addSlide();
      const text = slide.addRichText([{ runs: underlineRuns() }], {
        name: 'underline-text-owner',
      });
      const cellTable = slide.addTable([[{
        text: [{ runs: underlineRuns() }],
      }]], { name: 'underline-cell-owner' });
      const tableOwner = slide.addTable([[{
        text: [{ runs: underlineRuns() }],
      }]], { name: 'underline-table-owner-equivalent' });

      for (const paragraphs of [
        text.richText,
        placeholder.richText,
        cellTable.rows[0]!.cells[0]!.richText,
        tableOwner.rows[0]!.cells[0]!.richText,
      ]) {
        const snapshot = underlineSnapshot(paragraphs);
        expect(snapshot.slice(0, commonStyles.length).map((underline) =>
          typeof underline === 'object' ? underline.style : underline)).toEqual(commonStyles);
        expect(snapshot.slice(commonStyles.length)).toEqual([
          { style: 'sng' },
          false,
          { style: 'dbl', color: { kind: 'srgb', value: 'FF0000' } },
          { style: 'sng', color: { kind: 'scheme', value: 'accent2' } },
          { style: 'words' },
          { style: 'dotDashHeavy' },
        ]);
      }

      const detached = text.richText as unknown as Array<{
        runs: Array<{ style?: { underline?: { style?: string; color?: { value: string } } } }>;
      }>;
      detached[0]!.runs[0]!.style!.underline!.style = 'wavyDbl';
      detached[0]!.runs[commonStyles.length + 2]!.style!.underline!.color!.value = '00FF00';
      expect(underlineSnapshot(text.richText)[0]).toEqual({ style: 'dash' });
      expect(underlineSnapshot(text.richText)[commonStyles.length + 2]).toEqual({
        style: 'dbl',
        color: { kind: 'srgb', value: 'FF0000' },
      });

      const initialSlideXml = new TextDecoder().decode(
        created.opcPackage.requirePart(slide.partUri).bytes,
      );
      const initialLayoutXml = new TextDecoder().decode(
        created.opcPackage.requirePart(layout.partUri).bytes,
      );
      for (const [name, ownerTag] of [
        ['underline-text-owner', 'p:sp'],
        ['underline-cell-owner', 'p:graphicFrame'],
        ['underline-table-owner-equivalent', 'p:graphicFrame'],
      ] as const) {
        const ownerXml = namedOwnerXml(initialSlideXml, name, ownerTag);
        expect(underlineTokens(ownerXml)).toEqual(expectedTokens);
        expect(ownerXml).toContain(
          '<a:uFill><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:uFill>',
        );
        expect(ownerXml).toContain(
          '<a:uFill><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></a:uFill>',
        );
        expect(ownerXml).not.toContain('dotDashHeave"');
      }
      const placeholderXml = namedOwnerXml(
        initialLayoutXml,
        'underline-placeholder-owner',
        'p:sp',
      );
      expect(underlineTokens(placeholderXml)).toEqual(expectedTokens);
      expect(placeholderXml).not.toContain('dotDashHeave"');

      const beforeInvalid = created.opcPackage.requirePart(slide.partUri).bytes.slice();
      const beforeInvalidJournal = [...created.opcPackage.mutations];
      expect(() => {
        text.richText = [{ runs: [{
          text: 'Invalid upstream spelling',
          style: { underline: { style: 'dotDashHeave' as never } },
        }] }];
      }).toThrow(/underline/u);
      expect(created.opcPackage.requirePart(slide.partUri).bytes).toEqual(beforeInvalid);
      expect(created.opcPackage.mutations).toEqual(beforeInvalidJournal);

      created.duplicateSlide(0);
      text.richText = [{ runs: [
        { text: 'Disabled', style: { underline: false } },
        { text: 'Cleared' },
      ] }];
      placeholder.richText = [{ runs: [
        { text: 'Disabled', style: { underline: false } },
        { text: 'Cleared' },
      ] }];
      cellTable.setCellRichText(0, 0, [{ runs: [
        { text: 'Disabled', style: { underline: false } },
        { text: 'Cleared' },
      ] }]);
      tableOwner.setCellRichText(0, 0, [{ runs: [
        { text: 'Disabled', style: { underline: false } },
        { text: 'Cleared' },
      ] }]);

      const reopened = await PptxDocument.open(await created.write());
      const reopenedText = reopened.slides[0]!.shapes.find(
        (shape): shape is ShapeModel => shape instanceof ShapeModel
          && shape.name === 'underline-text-owner',
      )!;
      const reopenedCellTable = reopened.slides[0]!.shapes.find(
        (shape): shape is TableModel => shape instanceof TableModel
          && shape.name === 'underline-cell-owner',
      )!;
      const reopenedTableOwner = reopened.slides[0]!.shapes.find(
        (shape): shape is TableModel => shape instanceof TableModel
          && shape.name === 'underline-table-owner-equivalent',
      )!;
      const reopenedPlaceholder = reopened.layouts[0]!.placeholders.find(
        (shape): shape is ShapeModel => shape instanceof ShapeModel
          && shape.name === 'underline-placeholder-owner',
      )!;
      for (const paragraphs of [
        reopenedText.richText,
        reopenedPlaceholder.richText,
        reopenedCellTable.rows[0]!.cells[0]!.richText,
        reopenedTableOwner.rows[0]!.cells[0]!.richText,
      ]) {
        expect(underlineSnapshot(paragraphs)).toEqual([false, undefined]);
      }

      const duplicateText = reopened.slides[1]!.shapes.find(
        (shape): shape is ShapeModel => shape instanceof ShapeModel
          && shape.name === 'underline-text-owner',
      )!;
      expect(underlineSnapshot(duplicateText.richText).slice(0, commonStyles.length)
        .map((underline) => typeof underline === 'object' ? underline.style : underline))
        .toEqual(commonStyles);
      expect(underlineSnapshot(duplicateText.richText).at(-1)).toEqual({
        style: 'dotDashHeavy',
      });
      expect(reopened.diagnostics.filter(({ severity }) => severity === 'error')).toEqual([]);
      expect(validatePackage(reopened.opcPackage).filter(
        ({ severity }) => severity === 'error',
      )).toEqual([]);
    }
  });

  it('creates and reopens text direction owners in all six formats', async () => {
    const textDirections = [
      'eaVert',
      'horz',
      'mongolianVert',
      'vert',
      'vert270',
      'wordArtVert',
      'wordArtVertRtl',
    ] as const;
    const tableDirections = [
      'horz',
      'vert',
      'vert270',
      'wordArtVert',
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

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({ format });
      const slide = created.addSlide();
      const textShapes = textDirections.map((vert) =>
        slide.addText(`Text ${vert}`, {
          name: `text-direction-${vert}`,
          vert,
        }));
      const cellTable = slide.addTable([tableDirections.map((textDirection) => ({
        text: `Cell ${textDirection}`,
        options: { textDirection },
      }))], {
        name: 'text-direction-cell-owner',
      });
      const tableOwner = slide.addTable([['Inherited A', 'Inherited B']], {
        name: 'text-direction-table-owner',
        textDirection: 'vert270',
      });

      expect(textShapes.map(({ textDirection }) => textDirection)).toEqual(textDirections);
      expect(cellTable.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
        undefined,
        'vert',
        'vert270',
        'wordArtVert',
      ]);
      expect(tableOwner.textDirection).toBe('vert270');
      expect(tableOwner.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual([
        'vert270',
        'vert270',
      ]);

      const initialXml = new TextDecoder().decode(
        created.opcPackage.requirePart(slide.partUri).bytes,
      );
      expect(textDirections.map((vert) => bodyDirection(namedOwnerXml(
        initialXml,
        `text-direction-${vert}`,
        'p:sp',
      )))).toEqual(textDirections);
      expect(cellDirections(namedOwnerXml(
        initialXml,
        'text-direction-cell-owner',
        'p:graphicFrame',
      ))).toEqual([undefined, 'vert', 'vert270', 'wordArtVert']);
      expect(cellDirections(namedOwnerXml(
        initialXml,
        'text-direction-table-owner',
        'p:graphicFrame',
      ))).toEqual(['vert270', 'vert270']);
      expect(initialXml).not.toMatch(/<a:tcPr\b[^>]*\bvert="(?:eaVert|mongolianVert|wordArtVertRtl)"/u);

      const reopened = await PptxDocument.open(await created.write());
      const reopenedTextShapes = textDirections.map((vert) =>
        reopened.slides[0]!.shapes.find(
          (shape): shape is ShapeModel => shape instanceof ShapeModel
            && shape.name === `text-direction-${vert}`,
        )!);
      const reopenedCellTable = reopened.slides[0]!.shapes.find(
        (shape): shape is TableModel => shape instanceof TableModel
          && shape.name === 'text-direction-cell-owner',
      )!;
      const reopenedTableOwner = reopened.slides[0]!.shapes.find(
        (shape): shape is TableModel => shape instanceof TableModel
          && shape.name === 'text-direction-table-owner',
      )!;
      expect(reopenedTextShapes.map(({ textDirection }) => textDirection))
        .toEqual(textDirections);
      expect(reopenedCellTable.rows[0]!.cells.map(({ textDirection }) => textDirection))
        .toEqual([undefined, 'vert', 'vert270', 'wordArtVert']);
      expect(reopenedTableOwner.textDirection).toBe('vert270');
      expect(reopenedTableOwner.rows[0]!.cells.map(({ textDirection }) => textDirection))
        .toEqual(['vert270', 'vert270']);
      expect(reopened.diagnostics.filter(
        ({ severity }) => severity === 'error' || severity === 'warning',
      )).toEqual([]);
      expect(validatePackage(reopened.opcPackage).filter(
        ({ severity }) => severity === 'error' || severity === 'warning',
      )).toEqual([]);
    }
  });

  it('creates and reopens scalar text formatting owners in all six formats', async () => {
    const ownerXml = (xml: string, name: string, tag: 'p:sp' | 'p:graphicFrame') => {
      const nameOffset = xml.indexOf(`name="${name}"`);
      expect(nameOffset).toBeGreaterThanOrEqual(0);
      const start = xml.lastIndexOf(`<${tag}`, nameOffset);
      const end = xml.indexOf(`</${tag}>`, nameOffset);
      expect(start).toBeGreaterThanOrEqual(0);
      expect(end).toBeGreaterThan(nameOffset);
      return xml.slice(start, end + tag.length + 3);
    };
    const scalarStyle = {
      fontFamily: 'Aptos Display',
      fontSize: 17.25,
      lang: 'zh-CN',
      bold: true,
      italic: true,
      color: { kind: 'srgb' as const, value: 'C00000' },
      highlight: { kind: 'scheme' as const, value: 'accent2' },
    };
    const scalarRuns = (prefix: string) => [{
      runs: [
        { text: `${prefix} styled`, style: scalarStyle },
        {
          text: `${prefix} soft`,
          softBreakBefore: true,
          style: {
            fontFamily: 'Courier New',
            fontSize: 9.5,
            lang: 'fr-CA',
            bold: false,
            italic: false,
            color: { kind: 'scheme' as const, value: 'accent1' },
            highlight: { kind: 'srgb' as const, value: 'FFFF00' },
          },
        },
        { text: `${prefix} paragraph`, breakLine: true },
        { text: `${prefix} tail` },
      ],
    }];
    const scalarSnapshot = (paragraphs: readonly RichTextParagraph[]) => ({
      paragraphCount: paragraphs.length,
      first: paragraphs[0]!.runs[0]!.style,
      second: paragraphs[0]!.runs[1],
      tail: paragraphs.at(-1)!.runs[0]!.text,
    });

    for (const format of Object.keys(PRESENTATION_FORMAT_PROFILES) as PresentationFormat[]) {
      const created = PptxDocument.create({ format });
      const layout = created.layouts[0]!;
      const placeholder = layout.addPlaceholder(scalarRuns('Placeholder'), {
        name: 'scalar-placeholder-owner',
        type: 'body',
        index: 311,
      });
      const slide = created.addSlide();
      const text = slide.addRichText(scalarRuns('Text'), {
        name: 'scalar-text-owner',
      });
      const table = slide.addTable([[
        'Inherited defaults',
        {
          text: scalarRuns('Cell'),
          options: {
            fontFamily: 'Courier New',
            fontSize: 10,
            bold: false,
            color: { kind: 'srgb', value: '00AA00' },
          },
        },
      ]], {
        name: 'scalar-table-owner',
        fontFamily: 'Aptos',
        fontSize: 18.25,
        bold: true,
        color: { kind: 'scheme', value: 'accent1' },
      });
      slide.slideNumber = {
        style: {
          fontFamily: 'Aptos Narrow',
          fontSize: 14.5,
          lang: 'de-DE',
          bold: true,
          italic: true,
          color: { kind: 'scheme', value: 'accent3' },
        },
      };

      for (const paragraphs of [text.richText, placeholder.richText]) {
        expect(scalarSnapshot(paragraphs)).toMatchObject({
          paragraphCount: 2,
          first: scalarStyle,
          second: {
            text: expect.stringContaining('soft'),
            softBreakBefore: true,
            style: {
              fontFamily: 'Courier New',
              fontSize: 9.5,
              lang: 'fr-CA',
              bold: false,
              italic: false,
              color: { kind: 'scheme', value: 'accent1' },
              highlight: { kind: 'srgb', value: 'FFFF00' },
            },
          },
          tail: expect.stringContaining('tail'),
        });
      }
      expect(table.rows[0]!.cells[0]!.richText[0]!.runs[0]!.style).toMatchObject({
        fontFamily: 'Aptos',
        fontSize: 18.25,
        bold: true,
        color: { kind: 'scheme', value: 'accent1' },
      });
      expect(table.rows[0]!.cells[1]!.richText[0]!.runs[0]!.style).toMatchObject(scalarStyle);
      expect(table.rows[0]!.cells[1]!.richText[0]!.runs[1]).toMatchObject({
        softBreakBefore: true,
        style: { bold: false, italic: false },
      });
      expect(table.rows[0]!.cells[1]!.richText).toHaveLength(2);
      expect(slide.slideNumber?.style).toEqual({
        fontFamily: 'Aptos Narrow',
        fontSize: 14.5,
        lang: 'de-DE',
        bold: true,
        italic: true,
        color: { kind: 'scheme', value: 'accent3' },
      });

      const slideXml = new TextDecoder().decode(
        created.opcPackage.requirePart(slide.partUri).bytes,
      );
      const textXml = ownerXml(slideXml, 'scalar-text-owner', 'p:sp');
      const tableXml = ownerXml(slideXml, 'scalar-table-owner', 'p:graphicFrame');
      const layoutXml = new TextDecoder().decode(
        created.opcPackage.requirePart(layout.partUri).bytes,
      );
      const placeholderXml = ownerXml(layoutXml, 'scalar-placeholder-owner', 'p:sp');
      for (const xml of [textXml, tableXml, placeholderXml]) {
        expect(xml).toContain('lang="zh-CN"');
        expect(xml).toContain('sz="1725"');
        expect(xml).toContain('b="1" i="1"');
        expect(xml).toContain('<a:srgbClr val="C00000"/>');
        expect(xml).toContain('<a:highlight><a:schemeClr val="accent2"/></a:highlight>');
        expect(xml).toContain('<a:latin typeface="Aptos Display"/>');
        expect(xml).toContain('<a:br/>');
      }
      expect(tableXml).toContain('sz="1825"');
      expect(tableXml).toContain('sz="950"');
      expect(tableXml).toContain('b="0" i="0"');
      expect(slideXml).toContain('type="slidenum"');
      expect(slideXml).toContain('lang="de-DE"');
      expect(slideXml).toContain('sz="1450"');
      expect(slideXml).toContain('<a:schemeClr val="accent3"/>');

      const reopened = await PptxDocument.open(await created.write());
      const reopenedText = reopened.slides[0]!.shapes.find(
        (shape): shape is ShapeModel => shape instanceof ShapeModel
          && shape.name === 'scalar-text-owner',
      )!;
      const reopenedTable = reopened.slides[0]!.shapes.find(
        (shape): shape is TableModel => shape instanceof TableModel
          && shape.name === 'scalar-table-owner',
      )!;
      const reopenedPlaceholder = reopened.layouts[0]!.placeholders.find(
        (shape): shape is ShapeModel => shape instanceof ShapeModel
          && shape.name === 'scalar-placeholder-owner',
      )!;
      expect(scalarSnapshot(reopenedText.richText)).toEqual(scalarSnapshot(text.richText));
      expect(scalarSnapshot(reopenedPlaceholder.richText))
        .toEqual(scalarSnapshot(placeholder.richText));
      expect(reopenedTable.rows[0]!.cells.map(({ richText }) => richText))
        .toEqual(table.rows[0]!.cells.map(({ richText }) => richText));
      expect(reopened.slides[0]!.slideNumber).toEqual(slide.slideNumber);
      expect(reopened.diagnostics.filter(
        ({ severity }) => severity === 'error' || severity === 'warning',
      )).toEqual([]);
      expect(validatePackage(reopened.opcPackage).filter(
        ({ severity }) => severity === 'error' || severity === 'warning',
      )).toEqual([]);
    }
  });

  it('rejects malformed rich text values before changing the slide package state', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const shape = slide.addText('Original');
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    const invalid = [
      [],
      [null],
      [{}],
      [{ runs: null }],
      [{ runs: [null] }],
      [{ runs: [{ text: 42 }] }],
      [{ runs: [{ text: 'two\nlines' }] }],
      [{ runs: [{ text: 'invalid\u0000xml' }] }],
      [{ runs: [{ text: 'x', softBreakBefore: 'yes' }] }],
      [{ runs: [{ text: 'x', style: null }] }],
      [{ runs: [{ text: 'x', style: { fontFamily: '' } }] }],
      [{ runs: [{ text: 'x', style: { fontSize: 0 } }] }],
      [{ runs: [{ text: 'x', style: { fontSize: Number.POSITIVE_INFINITY } }] }],
      [{ runs: [{ text: 'x', style: { bold: 'yes' } }] }],
      [{ runs: [{ text: 'x', style: { color: { kind: 'srgb', value: 'red' } } }] }],
      [{ runs: [{ text: 'x', style: { color: { kind: 'scheme', value: 'unknown' } } }] }],
      [{ runs: [], align: 'middle' }],
      [{ runs: [], bullet: null }],
      [{ runs: [], bullet: {} }],
      [{ runs: [], bullet: { kind: 'bullet', character: '' } }],
      [{ runs: [], bullet: { kind: 'bullet', character: 'ab' } }],
      [{ runs: [], bullet: { kind: 'bullet', character: '\n' } }],
      [{ runs: [], bullet: { kind: 'bullet', indent: -1 } }],
      [{ runs: [], bullet: { kind: 'bullet', indent: 4033 } }],
      [{ runs: [], bullet: { kind: 'number', style: 'unsupported' } }],
      [{ runs: [], bullet: { kind: 'number', startAt: 0 } }],
      [{ runs: [], bullet: { kind: 'number', startAt: 1.5 } }],
      [{ runs: [], bullet: { kind: 'number', startAt: 32768 } }],
      [{ runs: [], bullet: { kind: 'number', color: 'red' } }],
      [{ runs: [], spacing: null }],
      [{ runs: [], spacing: {} }],
      [{ runs: [], spacing: { before: -1 } }],
      [{ runs: [], spacing: { before: 1584.01 } }],
      [{ runs: [], spacing: { after: Number.POSITIVE_INFINITY } }],
      [{ runs: [], spacing: { line: null } }],
      [{ runs: [], spacing: { line: {} } }],
      [{ runs: [], spacing: { line: { kind: 'exact', points: 0 } } }],
      [{ runs: [], spacing: { line: { kind: 'exact', points: 1585 } } }],
      [{ runs: [], spacing: { line: { kind: 'multiple', factor: 0 } } }],
      [{ runs: [], spacing: { line: { kind: 'multiple', factor: 133 } } }],
      [{ runs: [], spacing: { line: { kind: 'multiple', factor: 1.5, points: 12 } } }],
      [{ runs: [], spacing: { before: 2, unknown: true } }],
      [{ runs: [], level: null }],
      [{ runs: [], level: '2' }],
      [{ runs: [], level: Number.NaN }],
      [{ runs: [], level: -1 }],
      [{ runs: [], level: 1.5 }],
      [{ runs: [], level: 9 }],
      [{ runs: [], tabStops: null }],
      [{ runs: [], tabStops: {} }],
      [{ runs: [], tabStops: [null] }],
      [{ runs: [], tabStops: [{}] }],
      [{ runs: [], tabStops: [{ position: '1' }] }],
      [{ runs: [], tabStops: [{ position: Number.NaN }] }],
      [{ runs: [], tabStops: [{ position: Number.POSITIVE_INFINITY }] }],
      [{ runs: [], tabStops: [{ position: 3_000 }] }],
      [{ runs: [], tabStops: [{ position: 1, alignment: 'tab' }] }],
      [{ runs: [], tabStops: [{ position: 1, leader: 'dot' }] }],
      [{ runs: [{ text: 'x', style: { underline: null } }] }],
      [{ runs: [{ text: 'x', style: { underline: 'sng' } }] }],
      [{ runs: [{ text: 'x', style: { underline: [] } }] }],
      [{ runs: [{ text: 'x', style: { underline: {} } }] }],
      [{ runs: [{ text: 'x', style: { underline: { style: 'none' } } }] }],
      [{ runs: [{ text: 'x', style: { underline: { style: 'dotDashHeave' } } }] }],
      [{ runs: [{ text: 'x', style: { underline: { style: 'unknown' } } }] }],
      [{ runs: [{ text: 'x', style: { underline: { style: 1 } } }] }],
      [{ runs: [{ text: 'x', style: { underline: { color: null } } }] }],
      [{ runs: [{ text: 'x', style: { underline: { color: { kind: 'srgb', value: 'red' } } } }] }],
      [{ runs: [{ text: 'x', style: { underline: { style: 'sng', width: 2 } } }] }],
      [{ runs: [{ text: 'x', style: { strike: null } }] }],
      [{ runs: [{ text: 'x', style: { strike: 1 } }] }],
      [{ runs: [{ text: 'x', style: { strike: {} } }] }],
      [{ runs: [{ text: 'x', style: { strike: [] } }] }],
      [{ runs: [{ text: 'x', style: { strike: 'noStrike' } }] }],
      [{ runs: [{ text: 'x', style: { strike: 'tripleStrike' } }] }],
      [{ runs: [{ text: 'x', style: { highlight: null } }] }],
      [{ runs: [{ text: 'x', style: { highlight: true } }] }],
      [{ runs: [{ text: 'x', style: { highlight: 'FFFF00' } }] }],
      [{ runs: [{ text: 'x', style: { highlight: [] } }] }],
      [{ runs: [{ text: 'x', style: { highlight: {} } }] }],
      [{ runs: [{ text: 'x', style: { highlight: { kind: 'srgb', value: 'yellow' } } }] }],
      [{ runs: [{ text: 'x', style: { highlight: { kind: 'scheme', value: 'unknown' } } }] }],
      [{ runs: [{ text: 'x', style: { highlight: { kind: 'srgb', value: 'FFFF00', alpha: 1 } } }] }],
      [{ runs: [{ text: 'x', style: { baseline: null } }] }],
      [{ runs: [{ text: 'x', style: { baseline: true } }] }],
      [{ runs: [{ text: 'x', style: { baseline: false } }] }],
      [{ runs: [{ text: 'x', style: { baseline: {} } }] }],
      [{ runs: [{ text: 'x', style: { baseline: [] } }] }],
      [{ runs: [{ text: 'x', style: { baseline: 'super' } }] }],
      [{ runs: [{ text: 'x', style: { baseline: Number.NaN } }] }],
      [{ runs: [{ text: 'x', style: { baseline: Number.POSITIVE_INFINITY } }] }],
      [{ runs: [{ text: 'x', style: { baseline: -2_147_483.649 } }] }],
      [{ runs: [{ text: 'x', style: { baseline: 2_147_483.648 } }] }],
      [{ runs: [{ text: 'x', style: { characterSpacing: null } }] }],
      [{ runs: [{ text: 'x', style: { characterSpacing: true } }] }],
      [{ runs: [{ text: 'x', style: { characterSpacing: '1' } }] }],
      [{ runs: [{ text: 'x', style: { characterSpacing: {} } }] }],
      [{ runs: [{ text: 'x', style: { characterSpacing: [] } }] }],
      [{ runs: [{ text: 'x', style: { characterSpacing: Number.NaN } }] }],
      [{ runs: [{ text: 'x', style: { characterSpacing: Number.POSITIVE_INFINITY } }] }],
      [{ runs: [{ text: 'x', style: { characterSpacing: -21_474_836.49 } }] }],
      [{ runs: [{ text: 'x', style: { characterSpacing: 21_474_836.48 } }] }],
      [{ runs: [{ text: 'x', style: { glow: null } }] }],
      [{ runs: [{ text: 'x', style: { glow: true } }] }],
      [{ runs: [{ text: 'x', style: { glow: 'red' } }] }],
      [{ runs: [{ text: 'x', style: { glow: [] } }] }],
      [{ runs: [{ text: 'x', style: { glow: {} } }] }],
      [{ runs: [{ text: 'x', style: { glow: { size: 1 } } }] }],
      [{ runs: [{ text: 'x', style: { glow: { opacity: 0.5 } } }] }],
      [{ runs: [{ text: 'x', style: { glow: { opacity: 0.5, size: '1' } } }] }],
      [{ runs: [{ text: 'x', style: { glow: { opacity: '0.5', size: 1 } } }] }],
      [{ runs: [{ text: 'x', style: { glow: { opacity: 0.5, size: Number.NaN } } }] }],
      [{ runs: [{ text: 'x', style: { glow: { opacity: Number.POSITIVE_INFINITY, size: 1 } } }] }],
      [{ runs: [{ text: 'x', style: { glow: { opacity: 0.5, size: -0.01 } } }] }],
      [{ runs: [{ text: 'x', style: { glow: { opacity: 0.5, size: 2_147_483_648 } } }] }],
      [{ runs: [{ text: 'x', style: { glow: { opacity: -0.01, size: 1 } } }] }],
      [{ runs: [{ text: 'x', style: { glow: { opacity: 1.01, size: 1 } } }] }],
      [{ runs: [{ text: 'x', style: { glow: { color: { kind: 'srgb', value: 'red' }, opacity: 0.5, size: 1 } } }] }],
      [{ runs: [{ text: 'x', style: { glow: { color: { kind: 'srgb', value: 'FF0000' }, opacity: 0.5, size: 1, blur: 2 } } }] }],
      [{ runs: [{ text: 'x', style: { outline: null } }] }],
      [{ runs: [{ text: 'x', style: { outline: true } }] }],
      [{ runs: [{ text: 'x', style: { outline: 'red' } }] }],
      [{ runs: [{ text: 'x', style: { outline: [] } }] }],
      [{ runs: [{ text: 'x', style: { outline: {} } }] }],
      [{ runs: [{ text: 'x', style: { outline: { color: { kind: 'srgb', value: 'FF0000' } } } }] }],
      [{ runs: [{ text: 'x', style: { outline: { size: 1 } } }] }],
      [{ runs: [{ text: 'x', style: { outline: { color: { kind: 'srgb', value: 'FF0000' }, size: '1' } } }] }],
      [{ runs: [{ text: 'x', style: { outline: { color: { kind: 'srgb', value: 'FF0000' }, size: Number.NaN } } }] }],
      [{ runs: [{ text: 'x', style: { outline: { color: { kind: 'srgb', value: 'FF0000' }, size: Number.POSITIVE_INFINITY } } }] }],
      [{ runs: [{ text: 'x', style: { outline: { color: { kind: 'srgb', value: 'FF0000' }, size: -0.01 } } }] }],
      [{ runs: [{ text: 'x', style: { outline: { color: { kind: 'srgb', value: 'FF0000' }, size: 1584.1 } } }] }],
      [{ runs: [{ text: 'x', style: { outline: { color: { kind: 'srgb', value: 'red' }, size: 1 } } }] }],
      [{ runs: [{ text: 'x', style: { outline: { color: { kind: 'srgb', value: 'FF0000' }, size: 1, dash: 'solid' } } }] }],
      [{ runs: [{ text: 'x', style: { color: { kind: 'srgb', value: 'FF0000', alpha: 0.5 } } }] }],
    ];
    for (const value of invalid) {
      expect(() => slide.addRichText(value as never)).toThrow();
      expect(() => {
        shape.richText = value as never;
      }).toThrow();
    }
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
  });

  it('allocates text shape ids before extLst and rolls back rejected additions', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const part = document.opcPackage.requirePart(slide.partUri);
    const withExtension = new TextDecoder()
      .decode(part.bytes)
      .replace('</p:spTree>', '<p:extLst><p:ext uri="urn:test"><x:opaque xmlns:x="urn:test">KEEP</x:opaque></p:ext></p:extLst></p:spTree>');
    document.opcPackage.setPart(slide.partUri, withExtension, part.contentType);

    const first = slide.addText('First', { name: 'First' });
    const second = slide.addText('Second', { name: 'Second' });
    expect([first.id, second.id]).toEqual([2, 3]);
    const updatedXml = new TextDecoder().decode(document.opcPackage.requirePart(slide.partUri).bytes);
    expect(updatedXml.indexOf('name="Second"')).toBeLessThan(updatedXml.indexOf('<p:extLst>'));
    expect(updatedXml).toContain('<a:t xml:space="preserve">First</a:t>');
    expect(updatedXml).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');

    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];
    expect(() => {
      first.text = 'invalid\u0000xml';
    }).toThrow(/invalid XML characters/);
    expect(() => slide.addText('invalid\u0000xml')).toThrow(/invalid XML characters/);
    expect(() => slide.addText('bad width', { width: 0 as never })).toThrow(/width must be greater/);
    expect(() => slide.addText('bad coordinate', { x: Number.NaN as never })).toThrow(/x must be finite/);
    expect(() => slide.addText('bad flip', { flipHorizontal: 'yes' as never })).toThrow(/must be a boolean/);
    expect(() => slide.addText('bad alignment', { align: 'middle' as never })).toThrow(/left, center, right/);
    expect(() => slide.addText('bad bullet', { bullet: { kind: 'number', startAt: 0 } as never })).toThrow(/startAt/);
    expect(() => slide.addText('bad spacing', { spacing: false as never })).toThrow(/spacing/);
    expect(() => slide.addText('bad level', { level: 9 })).toThrow(/between 0 and 8/);
    expect(() => slide.addText('bad tabs', { tabStops: [{ position: 1, alignment: 'tab' as never }] })).toThrow(/alignment/);
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);

    let rolledBack: ShapeModel | undefined;
    expect(() =>
      document.transaction(() => {
        rolledBack = slide.addText('rollback');
        throw new Error('restore text shape');
      }),
    ).toThrow('restore text shape');
    expect(slide.shapes).toHaveLength(2);
    expect(() => rolledBack!.text).toThrow(ModelParseError);
  });

  it('does not mutate a malformed slide when its shape tree is missing', () => {
    const document = PptxDocument.create();
    const slide = document.addSlide();
    const part = document.opcPackage.requirePart(slide.partUri);
    document.opcPackage.setPart(
      slide.partUri,
      '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld/></p:sld>',
      part.contentType,
    );
    const before = document.opcPackage.requirePart(slide.partUri).bytes;
    const journal = [...document.opcPackage.mutations];

    expect(() => slide.addText('missing tree')).toThrow(/does not contain a shape tree/);
    expect(document.opcPackage.requirePart(slide.partUri).bytes).toEqual(before);
    expect(document.opcPackage.mutations).toEqual(journal);
  });

  it('opens Buffer-like values and streams, then returns unchanged bytes exactly', async () => {
    const input = await titleFixture();
    const arrayBuffer = input.slice().buffer as ArrayBuffer;
    const webStream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(input.slice(0, 20));
        controller.enqueue(input.slice(20));
        controller.close();
      },
    });
    const asyncBytes = {
      async *[Symbol.asyncIterator]() {
        yield input.slice(0, 30);
        yield input.slice(30);
      },
    };
    for (const source of [
      input,
      arrayBuffer,
      new Blob([new Uint8Array(input).buffer]),
      webStream,
      asyncBytes,
      Readable.from(input),
    ]) {
      const document = await PptxDocument.open(source);
      expect(document.codecRegistry.codecs.map(({ id }) => id)).toEqual([
        'builtin.master-layout-theme',
        'builtin.gradient-transparency',
        'builtin.media',
      ]);
      expect(document.masterLayoutTheme).toBe(document.masterLayoutTheme);
      expect(document.codecRegistry.codecs.find(({ id }) => id === 'builtin.master-layout-theme')).toBe(
        document.masterLayoutTheme,
      );
      expect(document.format).toBe('pptx');
      expect(document.slides[0]?.title.text).toBe('Original');
      expect(await document.write()).toEqual(input);
      const blob = await document.writeBlob();
      expect(blob.type).toBe('application/vnd.openxmlformats-officedocument.presentationml.presentation');
      expect(new Uint8Array(await blob.arrayBuffer())).toEqual(input);
    }
  });

  it('keeps Node path, stream, and writeFile adapters working', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'pptx-sdk-'));
    const inputPath = join(directory, 'input.pptx');
    const outputPath = join(directory, 'output.pptx');
    const input = await titleFixture();
    try {
      await writeFile(inputPath, input);
      const fromPath = await PptxDocument.open(inputPath);
      expect(await fromPath.write()).toEqual(input);
      const fromStream = await PptxDocument.open(openPptxStream(inputPath));
      await fromStream.writeFile(outputPath);
      expect(new Uint8Array(await readFile(outputPath))).toEqual(input);
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  });

  it('edits the title while preserving every untouched part payload', async () => {
    const input = await titleFixture();
    const before = await OpcPackage.open(input);
    const untouchedHashes = new Map(
      before.parts
        .filter(({ uri }) => uri !== '/ppt/slides/slide1.xml')
        .map(({ uri, bytes }) => [uri, hash(bytes)]),
    );
    const document = await PptxDocument.open(input);
    document.slides[0]!.title.text = 'Updated & preserved';
    const output = await document.write();
    const after = await OpcPackage.open(output);
    const slideXml = new TextDecoder().decode(after.requirePart('/ppt/slides/slide1.xml').bytes);
    expect(slideXml).toContain('<a:t xml:space="preserve">Updated &amp; preserved</a:t>');
    expect(slideXml).toContain('<x:unknown xmlns:x="x" custom="keep"/>');
    for (const [uri, expectedHash] of untouchedHashes) {
      expect(hash(after.requirePart(uri).bytes), uri).toBe(expectedHash);
    }
  });

  it('commits document transactions and rolls back package validation failures', async () => {
    const document = await PptxDocument.open(await titleFixture());
    const slideId = document.transaction((draft) => {
      draft.slides[0]!.title.text = 'Committed';
      return draft.addSlide().slideId;
    });
    expect(slideId).toBe(257);
    expect(document.slides.map(({ title }) => title.text)).toEqual(['Committed', '']);

    const rootRelationships = document.opcPackage.requirePart('/_rels/.rels').bytes;
    expect(() =>
      document.transaction((draft) => {
        draft.opcPackage.setPart(
          '/_rels/.rels',
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/missing.xml"/></Relationships>',
        );
      }),
    ).toThrow(ValidationError);
    expect(document.opcPackage.requirePart('/_rels/.rels').bytes).toEqual(rootRelationships);
    expect(document.slides.map(({ title }) => title.text)).toEqual(['Committed', '']);
  });
});

function hash(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}

const SDK_DEFAULT_POSTER_BYTES = Uint8Array.from([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1,
  0, 0, 0, 1, 8, 4, 0, 0, 0, 181, 28, 12, 2, 0, 0, 0, 11, 73, 68, 65, 84,
  120, 218, 99, 252, 255, 31, 0, 2, 235, 1, 245, 143, 89, 213, 153, 0, 0, 0,
  0, 73, 69, 78, 68, 174, 66, 96, 130,
]);

interface SdkMediaPictureState {
  readonly shapeId: number;
  readonly kind: MediaKind;
  readonly name: string | undefined;
  readonly altText: string | undefined;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly hasOfficeMedia: boolean;
  readonly relationshipTypes: readonly string[];
  readonly kindTargetMode: 'Internal' | 'External' | undefined;
  readonly kindTarget: string | undefined;
  readonly kindResolvedTarget: string | undefined;
}

function sdkMediaPictureStates(
  document: PptxDocument,
  slideIndex: number,
): SdkMediaPictureState[] {
  const slide = document.slides[slideIndex]!;
  const xml = LosslessXmlDocument.parse(document.opcPackage.requirePart(slide.partUri).bytes);
  const relationships = document.opcPackage.relationships(slide.partUri);
  return xml.elements('pic').flatMap((picture) => {
    const kindElement = xml.descendants(picture).find(({ localName }) =>
      localName === 'audioFile' || localName === 'videoFile');
    if (!kindElement) return [];
    const properties = xml.descendants(picture, 'cNvPr')[0]!;
    const transform = xml.descendants(picture, 'xfrm')[0]!;
    const offset = directSdkElementChildren(transform, 'off')[0]!;
    const extent = directSdkElementChildren(transform, 'ext')[0]!;
    const relationshipIds = xml.descendants(picture)
      .flatMap(({ attributes }) => attributes)
      .filter(({ name, value }) => (name === 'r:link' || name === 'r:embed') && value.length > 0)
      .map(({ value }) => value);
    const kindRelationshipId = xml.attribute(kindElement, 'r:link')?.value;
    const kindRelationship = relationships.find(({ id }) => id === kindRelationshipId);
    return [{
      shapeId: Number(xml.attribute(properties, 'id')?.value),
      kind: kindElement.localName === 'audioFile' ? 'audio' : 'video',
      name: xml.attribute(properties, 'name')?.value,
      altText: xml.attribute(properties, 'descr')?.value,
      x: Number(xml.attribute(offset, 'x')?.value),
      y: Number(xml.attribute(offset, 'y')?.value),
      width: Number(xml.attribute(extent, 'cx')?.value),
      height: Number(xml.attribute(extent, 'cy')?.value),
      hasOfficeMedia: xml.descendants(picture, 'media').length === 1,
      relationshipTypes: relationships
        .filter(({ id }) => relationshipIds.includes(id))
        .map(({ type }) => type),
      kindTargetMode: kindRelationship?.targetMode,
      kindTarget: kindRelationship?.target,
      kindResolvedTarget: kindRelationship?.resolvedTarget,
    }];
  });
}

function directSdkElementChildren(element: XmlElement, localName: string): XmlElement[] {
  return element.children.filter(
    (child): child is XmlElement => child.type === 'element' && child.localName === localName,
  );
}

function sdkNativeMediaTimingState(
  document: PptxDocument,
  slideIndex: number,
  shapeId: number,
): {
  readonly kind: MediaKind;
  readonly settings: Required<MediaPlaybackSettings>;
  readonly commands: readonly string[];
} {
  const slide = document.slides[slideIndex]!;
  const xml = LosslessXmlDocument.parse(document.opcPackage.requirePart(slide.partUri).bytes);
  const mediaNode = xml.elements('cMediaNode').find((candidate) =>
    xml.descendants(candidate, 'spTgt').some(
      (target) => Number(xml.attribute(target, 'spid')?.value) === shapeId,
    ));
  if (!mediaNode || (mediaNode.parent?.localName !== 'audio' && mediaNode.parent?.localName !== 'video')) {
    throw new Error(`Native media timing for shape ${shapeId} was not found`);
  }
  const playCommand = xml.elements('cmd').find((command) =>
    xml.attribute(command, 'cmd')?.value === 'playFrom(0.0)'
    && xml.descendants(command, 'spTgt').some(
      (target) => Number(xml.attribute(target, 'spid')?.value) === shapeId,
    ));
  const effect = playCommand
    ? sdkAncestorWithAttribute(xml, playCommand, 'cTn', 'presetClass', 'mediacall')
    : undefined;
  const main = effect
    ? sdkAncestorWithAttribute(xml, effect, 'cTn', 'nodeType', 'mainSeq')
    : undefined;
  const mainList = main ? directSdkElementChildren(main, 'childTnLst')[0] : undefined;
  const container = mainList?.children.find(
    (child): child is XmlElement => child.type === 'element'
      && Boolean(effect && child.start <= effect.start && child.end >= effect.end),
  );
  const startNode = container ? directSdkElementChildren(container, 'cTn')[0] : undefined;
  const startList = startNode ? directSdkElementChildren(startNode, 'stCondLst')[0] : undefined;
  const start = startList ? directSdkElementChildren(startList, 'cond')[0] : undefined;
  if (!playCommand || !start) throw new Error(`Native media play timing for shape ${shapeId} is incomplete`);
  const commands = xml.elements('cmd').filter((command) =>
    xml.descendants(command, 'spTgt').some(
      (target) => Number(xml.attribute(target, 'spid')?.value) === shapeId,
    )).map((command) => xml.attribute(command, 'cmd')?.value ?? '');
  const mediaTime = directSdkElementChildren(mediaNode, 'cTn')[0]!;
  return {
    kind: mediaNode.parent.localName,
    settings: {
      play: xml.attribute(start, 'delay')?.value === '0' ? 'auto' : 'click',
      loop: xml.attribute(mediaTime, 'repeatCount')?.value === 'indefinite',
      hideWhenStopped: xml.attribute(mediaNode, 'showWhenStopped')?.value === '0',
      volume: Number(xml.attribute(mediaNode, 'vol')?.value ?? 100_000) / 100_000,
    },
    commands,
  };
}

function sdkSlideTimingState(
  document: PptxDocument,
  slideIndex: number,
): { readonly ids: readonly number[]; readonly targets: readonly number[] } {
  const slide = document.slides[slideIndex]!;
  const xml = LosslessXmlDocument.parse(document.opcPackage.requirePart(slide.partUri).bytes);
  return {
    ids: xml.elements('cTn').map((node) => Number(xml.attribute(node, 'id')?.value)),
    targets: xml.elements('spTgt').map((target) => Number(xml.attribute(target, 'spid')?.value)),
  };
}

function sdkAncestorWithAttribute(
  xml: LosslessXmlDocument,
  element: XmlElement,
  localName: string,
  attributeName: string,
  value: string,
): XmlElement | undefined {
  let current = element.parent;
  while (current) {
    if (
      current.localName === localName
      && xml.attribute(current, attributeName)?.value === value
    ) return current;
    current = current.parent;
  }
  return undefined;
}

async function sdkPackageSnapshot(document: PptxDocument): Promise<unknown> {
  const pkg = document.opcPackage;
  const sources = pkg.parts
    .filter(({ uri }) => !uri.endsWith('.rels'))
    .map(({ uri }) => uri);
  return {
    parts: pkg.parts.map(({ uri, contentType, bytes }) => ({
      uri,
      contentType,
      bytes: new Uint8Array(bytes),
    })),
    relationships: ['/', ...sources].map((uri) => [
      uri,
      pkg.relationships(uri).map((relationship) => ({ ...relationship })),
    ]),
    graph: pkg.graph,
    output: new Uint8Array(await pkg.write()),
    journal: pkg.mutations.map((mutation) => ({ ...mutation })),
  };
}

function tableCellHorizontalAlignmentTokens(
  document: PptxDocument,
  slidePartUri: string,
): (string | undefined)[] {
  const xml = new TextDecoder().decode(
    document.opcPackage.requirePart(slidePartUri).bytes,
  );
  return [...xml.matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g)]
    .map((match) => match[0]!.match(/<a:pPr[^>]*\salgn="([^"]+)"/)?.[1]);
}
