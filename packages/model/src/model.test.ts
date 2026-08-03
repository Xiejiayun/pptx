import JSZip from 'jszip';
import { describe, expect, it, vi } from 'vitest';
import { MediaCodec, type SlideNumberOptions } from '@pptx/codecs';
import { LosslessXmlDocument } from '@pptx/lossless-xml';
import { OpcPackage, relativeRelationshipTarget, relationshipPartUri } from '@pptx/opc';
import {
  ChartModel,
  CustomGeometryEvaluationError,
  ImageModel,
  MediaModel,
  ModelParseError,
  PRESENTATION_FORMAT_PROFILES,
  PRESET_SHAPE_TYPES,
  PresentationModel,
  ShapeModel,
  TableModel,
  UnsupportedPresentationFormatError,
  degrees,
  emuToInches,
  inches,
  type AddTableCellOptions,
  type AddTableCellInput,
  type AddTableOptions,
  type CustomGeometry,
  type Hyperlink,
  type PresentationFormat,
  type RichTextParagraph,
  type ShapeArrows,
  type ShapeAdjustment,
  type ShapeArrowType,
  type ShapeFill,
  type ShapeLine,
  type ShapeLineDash,
  type ShapeShadow,
  type TableCellBorderInput,
  type TableCellBorders,
  type TableCellFill,
  type TableCellTextDirection,
  type TextAlignment,
  type TextBoxMarginInput,
  type TextBoxMargins,
  type TextBoxVerticalAlignment,
} from './index.js';
import { readShapeHyperlink } from './shape-hyperlink.internal.js';
import { readShapeAdjustments } from './shape-adjustments.internal.js';
import { readSimpleShadow } from './simple-shadow.internal.js';
import { readCustomGeometry } from './custom-geometry.internal.js';
import { chartWorkbookMatches } from './chart-workbook.internal.js';
import { resolveSlideLayoutPartUri } from './presentation-layout.internal.js';
import {
  normalizePlaceholderIdentity,
  normalizePlaceholderSelector,
  readShapePlaceholder,
} from './placeholder.internal.js';

const CORE_PROPERTIES_RELATIONSHIP =
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const CORE_PROPERTIES_CONTENT_TYPE =
  'application/vnd.openxmlformats-package.core-properties+xml';
const EXTENDED_PROPERTIES_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';
const EXTENDED_PROPERTIES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.extended-properties+xml';
const HYPERLINK_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const SLIDE_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
const IMAGE_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const SLIDE_LAYOUT_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout';
const SLIDE_MASTER_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster';
const SLIDE_LAYOUT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml';
const SLIDE_MASTER_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml';

const customTriangleGeometry: CustomGeometry = {
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

const customFormulaGeometry: CustomGeometry = {
  adjustments: [
    { name: 'adj1', formula: { operator: 'val', operands: [25_000] } },
    { name: 'adj2', formula: { operator: 'pin', operands: [0, 75_000, 100_000] } },
  ],
  guides: [
    { name: 'x1', formula: { operator: '*/', operands: ['w', 'adj1', 100_000] } },
    { name: 'y1', formula: { operator: '+-', operands: ['h', 0, 'x1'] } },
    { name: 'a1', formula: { operator: 'at2', operands: ['y1', 'x1'] } },
    { name: 'r1', formula: { operator: 'max', operands: ['x1', 1] } },
  ],
  paths: [{
    width: 100_000,
    height: 100_000,
    fill: 'norm',
    commands: [
      { kind: 'moveTo', point: { x: 'x1', y: 0 } },
      {
        kind: 'quadraticBezierTo',
        control: { x: 'wd2', y: 'y1' },
        end: { x: 'r', y: 'b' },
      },
      {
        kind: 'arcTo',
        widthRadius: 'r1',
        heightRadius: 'hd2',
        startAngle: 'a1',
        sweepAngle: 'cd2',
      },
      { kind: 'close' },
    ],
  }],
};

const customFormulaReplacement: CustomGeometry = {
  adjustments: [{ name: 'adj1', formula: { operator: 'val', operands: [50_000] } }],
  guides: [
    { name: 'x1', formula: { operator: '*/', operands: ['w', 'adj1', 100_000] } },
    { name: 'y1', formula: { operator: 'min', operands: ['h', 'x1'] } },
  ],
  paths: [{
    width: 100_000,
    height: 100_000,
    fill: 'none',
    commands: [
      { kind: 'moveTo', point: { x: 'x1', y: 'y1' } },
      { kind: 'lineTo', point: { x: 'r', y: 0 } },
      { kind: 'close' },
    ],
  }],
};

const customEvaluationGeometry: CustomGeometry = {
  adjustments: [{ name: 'adj', formula: { operator: 'val', operands: [25_000] } }],
  guides: [
    { name: 'x1', formula: { operator: '*/', operands: ['w', 'adj', 100_000] } },
    { name: 'y1', formula: { operator: '*/', operands: ['h', 'adj', 100_000] } },
    { name: 'rad', formula: { operator: 'max', operands: ['x1', 1] } },
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
      {
        kind: 'arcTo',
        widthRadius: 'rad',
        heightRadius: 'hd2',
        startAngle: 0,
        sweepAngle: 'cd4',
      },
      { kind: 'close' },
    ],
  }],
};

const customHandleGeometry: CustomGeometry = {
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

const customHandleReplacement: CustomGeometry = {
  ...customHandleGeometry,
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

const customConnectionGeometry: CustomGeometry = {
  ...customHandleGeometry,
  connectionSites: [
    { angle: 0, position: { x: 'hc', y: 't' } },
    { angle: 'adjAng', position: { x: 'r', y: 'adjY' } },
    { angle: -5_400_000, position: { x: 25_000, y: 100_000 } },
    { angle: 0, position: { x: 'hc', y: 't' } },
  ],
};

const customConnectionReplacement: CustomGeometry = {
  ...customConnectionGeometry,
  connectionSites: [
    { angle: -5_400_000, position: { x: 25_000, y: 100_000 } },
    { angle: 0, position: { x: 'hc', y: 't' } },
    { angle: 'adjAng', position: { x: 'l', y: 'adjY' } },
    { angle: 0, position: { x: 'hc', y: 't' } },
  ],
};

const customTextRectangleGeometry: CustomGeometry = {
  ...customConnectionGeometry,
  textRectangle: {
    left: 'x1',
    top: 10_000,
    right: 'r',
    bottom: 90_000,
  },
};

const customTextRectangleReplacement: CustomGeometry = {
  ...customConnectionReplacement,
  textRectangle: {
    left: 0,
    top: 't',
    right: 80_000,
    bottom: 'b',
  },
};

function readCreatedShapeHyperlink(
  model: PresentationModel,
  slide: ReturnType<PresentationModel['addSlide']>,
  shapeId: number,
) {
  const xml = LosslessXmlDocument.parse(model.opcPackage.requirePart(slide.partUri).bytes);
  const shape = xml.elements('sp').find((candidate) => {
    const properties = xml.descendants(candidate, 'cNvPr')[0];
    return properties && xml.attribute(properties, 'id')?.value === String(shapeId);
  });
  if (!shape) throw new Error(`Shape ${shapeId} was not found`);
  return readShapeHyperlink(xml, shape, {
    relationships: slide.relationships,
    slidePartUris: model.slides.map(({ partUri }) => partUri),
  });
}

function readCreatedShapeShadow(
  model: PresentationModel,
  slide: ReturnType<PresentationModel['addSlide']>,
  shapeId: number,
) {
  const xml = LosslessXmlDocument.parse(model.opcPackage.requirePart(slide.partUri).bytes);
  const shape = xml.elements('sp').find((candidate) => {
    const properties = xml.descendants(candidate, 'cNvPr')[0];
    return properties && xml.attribute(properties, 'id')?.value === String(shapeId);
  });
  const shadow = shape
    ? xml.descendants(shape).find(({ localName }) =>
        localName === 'outerShdw' || localName === 'innerShdw')
    : undefined;
  if (!shadow) return undefined;
  return readSimpleShadow(shadow, 'a:');
}

function readCreatedShapeAdjustments(
  model: PresentationModel,
  slide: ReturnType<PresentationModel['addSlide']>,
  shapeId: number,
) {
  const xml = LosslessXmlDocument.parse(model.opcPackage.requirePart(slide.partUri).bytes);
  const shape = xml.elements('sp').find((candidate) => {
    const properties = xml.descendants(candidate, 'cNvPr')[0];
    return properties && xml.attribute(properties, 'id')?.value === String(shapeId);
  });
  if (!shape) throw new Error(`Shape ${shapeId} was not found`);
  return readShapeAdjustments(xml, shape);
}

function readCreatedCustomGeometry(
  model: PresentationModel,
  slide: ReturnType<PresentationModel['addSlide']>,
  shapeId: number,
) {
  const xml = LosslessXmlDocument.parse(model.opcPackage.requirePart(slide.partUri).bytes);
  const shape = xml.elements('sp').find((candidate) => {
    const properties = xml.descendants(candidate, 'cNvPr')[0];
    return properties && xml.attribute(properties, 'id')?.value === String(shapeId);
  });
  if (!shape) throw new Error(`Shape ${shapeId} was not found`);
  return readCustomGeometry(xml, shape);
}

function packageSnapshot(pkg: OpcPackage) {
  return {
    parts: pkg.parts.map(({ uri, contentType, bytes, relationships }) => ({
      uri,
      contentType,
      bytes: bytes.slice(),
      relationships,
    })),
    graph: pkg.graph,
    mutations: [...pkg.mutations],
  };
}

function slideNumberCache(pkg: OpcPackage, slidePartUri: string): string | undefined {
  const xml = LosslessXmlDocument.parse(pkg.requirePart(slidePartUri).bytes);
  const field = xml.elements('fld').find(
    (candidate) => xml.attribute(candidate, 'type')?.value === 'slidenum',
  );
  const text = field ? xml.descendants(field, 't') : [];
  return text.length === 1 ? xml.text(text[0]!) : undefined;
}

function emptyPresentationModel(): { pkg: OpcPackage; model: PresentationModel } {
  const pkg = OpcPackage.create();
  pkg.transaction(() => {
    pkg.setPart(
      '/ppt/presentation.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
        '<p:sldIdLst/><p:sldSz cx="9144000" cy="5143500"/>' +
        '<p:notesSz cx="5143500" cy="9144000"/></p:presentation>',
      PRESENTATION_FORMAT_PROFILES.pptx.presentationContentType,
    );
    pkg.addRelationship('/', {
      id: 'rId1',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
      target: 'ppt/presentation.xml',
    });
  });
  return { pkg, model: new PresentationModel(pkg) };
}

async function modelFixture(
  presentationContentType = PRESENTATION_FORMAT_PROFILES.pptx.presentationContentType,
): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="${presentationContentType}"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/ppt/embeddings/workbook1.xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/><Override PartName="/ppt/custom/opaque1.bin" ContentType="application/octet-stream"/><Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/><Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/></Types>`);
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="257" r:id="rId2"/><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500" type="screen" custom="KEEP"/><p:notesSz cx="5143500" cy="9144000"/><x:unknown xmlns:x="urn:test"/></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="p" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="r" xmlns:c="c"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm rot="60000"><a:off x="914400" y="1828800"/><a:ext cx="2743200" cy="914400"/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:t>First title</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id="3" name="Image 1"/></p:nvPicPr><p:blipFill><a:blip r:embed="rId1"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="20"/></a:xfrm></p:spPr></p:pic><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>A1</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>B1</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Chart 1"/></p:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rId2"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>');
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/><Relationship Id="rId3" Type="urn:example:relationships/opaque" Target="../custom/opaque1.bin"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>');
  zip.file('ppt/slides/slide2.xml', '<p:sld xmlns:p="p" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 2"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Second title</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>');
  zip.file('ppt/media/image1.png', new Uint8Array([137, 80, 78, 71]));
  zip.file('ppt/charts/chart1.xml', '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="10"/><c:axId val="20"/></c:barChart><c:catAx><c:axId val="10"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:crossAx val="10"/></c:valAx></c:plotArea></c:chart><c:externalData r:id="rId1"/></c:chartSpace>');
  zip.file('ppt/charts/_rels/chart1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/workbook1.xlsx"/></Relationships>');
  zip.file('ppt/embeddings/workbook1.xlsx', new Uint8Array([80, 75, 3, 4, 1]));
  zip.file('ppt/custom/opaque1.bin', new Uint8Array([9, 8, 7]));
  zip.file('ppt/notesSlides/notesSlide1.xml', '<p:notes xmlns:p="p"><p:cSld><p:spTree/></p:cSld></p:notes>');
  zip.file('ppt/notesSlides/_rels/notesSlide1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/></Relationships>');
  zip.file('ppt/notesMasters/notesMaster1.xml', '<p:notesMaster xmlns:p="p"><p:cSld><p:spTree/></p:cSld></p:notesMaster>');
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

describe('PresentationModel', () => {
  it('normalizes and strictly reads placeholder identity', () => {
    const normalized = normalizePlaceholderIdentity({ type: 'title', index: 103 });
    expect(normalized).toEqual({ type: 'title', index: 103 });
    expect(Object.isFrozen(normalized)).toBe(true);
    expect(normalizePlaceholderIdentity({ type: 'body', index: 0 }))
      .toEqual({ type: 'body', index: 0 });
    expect(normalizePlaceholderIdentity({ type: 'media', index: 4_294_967_294 }))
      .toEqual({ type: 'media', index: 4_294_967_294 });
    expect(normalizePlaceholderSelector('title_box')).toBe('title_box');
    expect(normalizePlaceholderSelector({ type: 'body', index: 7 }))
      .toEqual({ type: 'body', index: 7 });
    expect(() => normalizePlaceholderSelector('')).toThrow();
    expect(() => normalizePlaceholderSelector(Object.defineProperty({}, 'type', {
      enumerable: true,
      get: () => 'title',
    }))).toThrow();

    const inherited = Object.create({ type: 'title', index: 1 });
    for (const invalid of [
      inherited,
      { type: 'ctrTitle', index: 1 },
      { type: 'title', index: -1 },
      { type: 'title', index: 4_294_967_295 },
      { type: 'title', index: 1.5 },
      { type: 'title', index: 1, extra: true },
    ]) {
      expect(() => normalizePlaceholderIdentity(invalid)).toThrow();
    }

    const alternate = LosslessXmlDocument.parse(
      '<q:sp xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main">'
        + '<q:nvSpPr><q:cNvPr id="2" name="Title"/><q:cNvSpPr/><q:nvPr>'
        + '<q:ph type="title" idx="103"/></q:nvPr></q:nvSpPr></q:sp>',
    );
    expect(readShapePlaceholder(alternate, alternate.roots[0]!)).toEqual({
      type: 'title',
      index: 103,
    });
    expect(Object.isFrozen(readShapePlaceholder(alternate, alternate.roots[0]!))).toBe(true);

    for (const placeholder of [
      '<q:ph xmlns:x="urn:wrong" x:type="title" idx="1"/>',
      '<q:ph type="title" type="body" idx="1"/>',
      '<q:ph type="ctrTitle" idx="1"/>',
      '<x:ph xmlns:x="urn:wrong" type="title" idx="1"/>',
    ]) {
      const xml = LosslessXmlDocument.parse(
        '<q:sp xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main">'
          + '<q:nvSpPr><q:cNvPr id="2" name="Title"/><q:cNvSpPr/><q:nvPr>'
          + `${placeholder}</q:nvPr></q:nvSpPr></q:sp>`,
      );
      expect(readShapePlaceholder(xml, xml.roots[0]!)).toBeUndefined();
    }
  });

  it('populate image chart placeholder owners in place through the model API', async () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const layoutPartUri = '/ppt/slideLayouts/slideLayout1.xml';
    const placeholder = (
      id: number,
      name: string,
      type: 'pic' | 'chart',
      index: number,
      x: number,
      y: number,
      width: number,
      height: number,
    ) => '<p:sp><p:nvSpPr>'
      + `<p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>`
      + `<p:ph type="${type}" idx="${index}"/></p:nvPr></p:nvSpPr><p:spPr>`
      + `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/>`
      + '</a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>';
    const pictureOwner = placeholder(
      2,
      'hero_image',
      'pic',
      104,
      inches(1),
      inches(2),
      inches(3),
      inches(4),
    );
    const chartOwner = placeholder(
      3,
      'revenue_chart',
      'chart',
      105,
      inches(5),
      inches(1),
      inches(2),
      inches(3),
    );
    pkg.setPart(
      layoutPartUri,
      '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        + `<p:cSld name="DEFAULT"><p:spTree>${pictureOwner}${chartOwner}`
        + '</p:spTree></p:cSld></p:sldLayout>',
      SLIDE_LAYOUT_CONTENT_TYPE,
    );
    pkg.addRelationship(slide.partUri, {
      type: SLIDE_LAYOUT_RELATIONSHIP,
      target: '../slideLayouts/slideLayout1.xml',
    });
    pkg.setPart(
      slide.partUri,
      '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        + `<p:cSld><p:spTree>${pictureOwner}${chartOwner}</p:spTree></p:cSld></p:sld>`,
      pkg.requirePart(slide.partUri).contentType,
    );
    const emptyPicture = slide.shapes[0]!;
    const emptyChart = slide.shapes[1]!;

    const picture = slide.addImage(Uint8Array.of(137, 80, 78, 71), {
      contentType: 'image/png',
      placeholder: 'hero_image',
      x: inches(9),
      width: inches(9),
      sourceRectangle: { left: 10, top: 5, right: 10, bottom: 5 },
    });
    const chart = await slide.addChart('bar', [{
      name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20],
    }], {
      placeholder: { type: 'chart', index: 105 },
      x: inches(9),
      width: inches(9),
    });

    expect(picture).toBeInstanceOf(ImageModel);
    expect(picture).toMatchObject({
      id: emptyPicture.id,
      name: 'hero_image',
      placeholder: { type: 'pic', index: 104 },
      transform: {
        x: inches(1),
        y: inches(2),
        width: inches(3),
        height: inches(4),
        rotation: 0,
        flipHorizontal: false,
        flipVertical: false,
      },
      sourceRectangle: { left: 10, top: 5, right: 10, bottom: 5 },
    });
    expect(chart).toBeInstanceOf(ChartModel);
    expect(chart).toMatchObject({
      id: emptyChart.id,
      name: 'revenue_chart',
      placeholder: { type: 'chart', index: 105 },
      transform: {
        x: inches(5),
        y: inches(1),
        width: inches(2),
        height: inches(3),
        rotation: 0,
        flipHorizontal: false,
        flipVertical: false,
      },
    });
    expect(() => emptyPicture.name).toThrow(/stale/i);
    expect(() => emptyChart.name).toThrow(/stale/i);
    expect(slide.shapes).toEqual([picture, chart]);
    expect(new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes)).toMatch(
      /<p:pic>[\s\S]*<p:ph type="pic" idx="104"\/>[\s\S]*<p:graphicFrame[\s\S]*<p:ph type="chart" idx="105"\/>/,
    );
    expect(pkg.parts.filter(({ uri }) => uri.startsWith('/ppt/media/'))).toHaveLength(1);
    expect(pkg.parts.filter(({ contentType }) =>
      contentType === 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'))
      .toHaveLength(1);
    expect(pkg.parts.filter(({ contentType }) =>
      contentType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))
      .toHaveLength(1);

    const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
    const duplicatePicture = duplicate.shapes[0] as ImageModel;
    const duplicateChart = duplicate.shapes[1] as ChartModel;
    expect(duplicatePicture.sourcePartUri).toBe(picture.sourcePartUri);
    expect(duplicateChart.chartPartUri).not.toBe(chart.chartPartUri);
    expect(duplicateChart.workbookPartUri).not.toBe(chart.workbookPartUri);
    duplicatePicture.replaceData(Uint8Array.of(1, 2, 3), 'image/png');
    expect(duplicatePicture.sourcePartUri).not.toBe(picture.sourcePartUri);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    expect(reopened.slides[0]?.shapes.map(({ placeholder: identity }) => identity)).toEqual([
      { type: 'pic', index: 104 },
      { type: 'chart', index: 105 },
    ]);

    const rollbackFixture = emptyPresentationModel();
    const rollbackSlide = rollbackFixture.model.addSlide();
    rollbackFixture.pkg.setPart(
      layoutPartUri,
      '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        + `<p:cSld><p:spTree>${pictureOwner}</p:spTree></p:cSld></p:sldLayout>`,
      SLIDE_LAYOUT_CONTENT_TYPE,
    );
    rollbackFixture.pkg.addRelationship(rollbackSlide.partUri, {
      type: SLIDE_LAYOUT_RELATIONSHIP,
      target: '../slideLayouts/slideLayout1.xml',
    });
    rollbackFixture.pkg.setPart(
      rollbackSlide.partUri,
      '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        + `<p:cSld><p:spTree>${pictureOwner}</p:spTree></p:cSld></p:sld>`,
      rollbackFixture.pkg.requirePart(rollbackSlide.partUri).contentType,
    );
    const beforeRollback = packageSnapshot(rollbackFixture.pkg);
    expect(() => rollbackFixture.pkg.transaction(() => {
      rollbackSlide.addImage(Uint8Array.of(1), {
        contentType: 'image/png',
        placeholder: 'hero_image',
      });
      throw new Error('rollback placeholder image');
    })).toThrow('rollback placeholder image');
    expect(packageSnapshot(rollbackFixture.pkg)).toEqual(beforeRollback);
  });

  it('populate table media placeholder owners in place through the model API', async () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const layoutPartUri = '/ppt/slideLayouts/slideLayout1.xml';
    const placeholder = (
      id: number,
      name: string,
      type: 'tbl' | 'media',
      index: number,
      x: number,
      y: number,
      width: number,
      height: number,
    ) => '<p:sp><p:nvSpPr>'
      + `<p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>`
      + `<p:ph type="${type}" idx="${index}"/></p:nvPr></p:nvSpPr><p:spPr>`
      + `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${width}" cy="${height}"/>`
      + '</a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>';
    const tableOwner = placeholder(
      2,
      'data_table',
      'tbl',
      106,
      inches(1),
      inches(2),
      inches(4),
      inches(2),
    );
    const mediaOwner = placeholder(
      3,
      'narration',
      'media',
      107,
      inches(5),
      inches(2),
      inches(2),
      inches(1),
    );
    pkg.setPart(
      layoutPartUri,
      '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        + `<p:cSld><p:spTree>${tableOwner}${mediaOwner}</p:spTree></p:cSld>`
        + '</p:sldLayout>',
      SLIDE_LAYOUT_CONTENT_TYPE,
    );
    pkg.addRelationship(slide.partUri, {
      type: SLIDE_LAYOUT_RELATIONSHIP,
      target: '../slideLayouts/slideLayout1.xml',
    });
    pkg.setPart(
      slide.partUri,
      '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        + 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        + 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        + `<p:cSld><p:spTree>${tableOwner}${mediaOwner}</p:spTree></p:cSld></p:sld>`,
      pkg.requirePart(slide.partUri).contentType,
    );
    const emptyTable = slide.shapes[0]!;
    const emptyMedia = slide.shapes[1]!;

    const table = slide.addTable([['A', 'B'], ['1', '2']], {
      placeholder: 'data_table',
      width: inches(6),
      height: inches(4),
      columnWidths: [inches(2), inches(4)],
      rowHeights: [inches(1), inches(3)],
    });
    const audio = await slide.addAudio(Uint8Array.of(1, 2, 3), {
      placeholder: { type: 'media', index: 107 },
      contentType: 'audio/mpeg',
      poster: Uint8Array.of(4),
      posterContentType: 'image/png',
      play: 'auto',
      x: inches(9),
      width: inches(9),
    });

    expect(table).toMatchObject({
      id: emptyTable.id,
      name: 'data_table',
      placeholder: { type: 'tbl', index: 106 },
      transform: {
        x: inches(1),
        y: inches(2),
        width: inches(4),
        height: inches(2),
        rotation: 0,
        flipHorizontal: false,
        flipVertical: false,
      },
    });
    expect(table.columnWidths).toEqual([inches(4) / 3, inches(8) / 3]);
    expect(table.rowHeights).toEqual([inches(0.5), inches(1.5)]);
    expect(audio).toMatchObject({
      id: emptyMedia.id,
      name: 'narration',
      placeholder: { type: 'media', index: 107 },
      transform: {
        x: inches(5),
        y: inches(2),
        width: inches(2),
        height: inches(1),
        rotation: 0,
        flipHorizontal: false,
        flipVertical: false,
      },
      settings: {
        play: 'auto',
        loop: false,
        hideWhenStopped: false,
        volume: 1,
      },
    });
    expect(() => emptyTable.name).toThrow(/stale/i);
    expect(() => emptyMedia.name).toThrow(/stale/i);
    expect(slide.shapes).toEqual([table, audio]);
    const source = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(source).toContain('<p:ph type="tbl" idx="106"/>');
    expect(source).toContain('<p:ph type="media" idx="107"/>');
    expect(source).toContain('<p:timing>');

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    expect(reopened.slides[0]?.shapes.map(({ placeholder: identity }) => identity)).toEqual([
      { type: 'tbl', index: 106 },
      { type: 'media', index: 107 },
    ]);
  });

  it('resolves named slide layouts strictly without package mutation', () => {
    const { pkg, model } = emptyPresentationModel();
    const masterPartUri = '/ppt/slideMasters/slideMaster1.xml';
    const defaultLayoutPartUri = '/ppt/slideLayouts/slideLayout1.xml';
    const brandLayoutPartUri = '/ppt/slideLayouts/slideLayout2.xml';
    pkg.transaction(() => {
      pkg.setPart(
        masterPartUri,
        '<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld/></p:sldMaster>',
        SLIDE_MASTER_CONTENT_TYPE,
      );
      pkg.setPart(
        defaultLayoutPartUri,
        '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="DEFAULT"/></p:sldLayout>',
        SLIDE_LAYOUT_CONTENT_TYPE,
      );
      pkg.setPart(
        brandLayoutPartUri,
        '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="BRAND"/></p:sldLayout>',
        SLIDE_LAYOUT_CONTENT_TYPE,
      );
      pkg.addRelationship(model.presentationPartUri, {
        type: SLIDE_MASTER_RELATIONSHIP,
        target: 'slideMasters/slideMaster1.xml',
      });
      pkg.addRelationship(masterPartUri, {
        type: SLIDE_LAYOUT_RELATIONSHIP,
        target: '../slideLayouts/slideLayout1.xml',
      });
      pkg.addRelationship(masterPartUri, {
        type: SLIDE_LAYOUT_RELATIONSHIP,
        target: '../slideLayouts/slideLayout2.xml',
      });
    });

    const wrongNamespacePartUri = '/ppt/slideLayouts/slideLayout-wrong-namespace.xml';
    const qualifiedNamePartUri = '/ppt/slideLayouts/slideLayout-qualified-name.xml';
    const wrongContentPartUri = '/ppt/slideLayouts/slideLayout-wrong-content.xml';
    pkg.transaction(() => {
      pkg.setPart(
        wrongNamespacePartUri,
        '<x:sldLayout xmlns:x="urn:not-presentation"><x:cSld name="WRONG_NAMESPACE"/></x:sldLayout>',
        SLIDE_LAYOUT_CONTENT_TYPE,
      );
      pkg.setPart(
        qualifiedNamePartUri,
        '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:not-presentation"><p:cSld x:name="QUALIFIED_NAME"/></p:sldLayout>',
        SLIDE_LAYOUT_CONTENT_TYPE,
      );
      pkg.setPart(
        wrongContentPartUri,
        '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="WRONG_CONTENT"/></p:sldLayout>',
        'application/xml',
      );
      for (const target of [
        '../slideLayouts/slideLayout-wrong-namespace.xml',
        '../slideLayouts/slideLayout-qualified-name.xml',
        '../slideLayouts/slideLayout-wrong-content.xml',
      ]) {
        pkg.addRelationship(masterPartUri, {
          type: SLIDE_LAYOUT_RELATIONSHIP,
          target,
        });
      }
      pkg.addRelationship(masterPartUri, {
        type: SLIDE_LAYOUT_RELATIONSHIP,
        target: 'https://example.com/external-layout.xml',
        targetMode: 'External',
      });
      const relationshipsPartUri = relationshipPartUri(masterPartUri);
      const relationshipsPart = pkg.requirePart(relationshipsPartUri);
      const relationshipsXml = new TextDecoder().decode(relationshipsPart.bytes).replace(
        '</Relationships>',
        `<Relationship Id="rIdDangling" Type="${SLIDE_LAYOUT_RELATIONSHIP}" Target="../slideLayouts/missing.xml"/></Relationships>`,
      );
      pkg.setPart(
        relationshipsPartUri,
        relationshipsXml,
        relationshipsPart.contentType,
      );
    });

    expect(resolveSlideLayoutPartUri(
      pkg,
      model.presentationPartUri,
      'BRAND',
      undefined,
    )).toBe(brandLayoutPartUri);
    expect(resolveSlideLayoutPartUri(
      pkg,
      model.presentationPartUri,
      undefined,
      brandLayoutPartUri,
    )).toBe(brandLayoutPartUri);
    expect(resolveSlideLayoutPartUri(
      pkg,
      model.presentationPartUri,
      undefined,
      undefined,
    )).toBe(defaultLayoutPartUri);

    const before = packageSnapshot(pkg);
    expect(() => resolveSlideLayoutPartUri(
      pkg,
      model.presentationPartUri,
      'MISSING',
      undefined,
    )).toThrow(RangeError);
    for (const ignoredName of [
      'WRONG_NAMESPACE',
      'QUALIFIED_NAME',
      'WRONG_CONTENT',
    ]) {
      expect(() => resolveSlideLayoutPartUri(
        pkg,
        model.presentationPartUri,
        ignoredName,
        undefined,
      )).toThrow(RangeError);
    }
    expect(packageSnapshot(pkg)).toEqual(before);

    const duplicatePartUri = '/ppt/slideLayouts/slideLayout3.xml';
    pkg.transaction(() => {
      pkg.setPart(
        duplicatePartUri,
        '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld name="BRAND"/></p:sldLayout>',
        SLIDE_LAYOUT_CONTENT_TYPE,
      );
      pkg.addRelationship(masterPartUri, {
        type: SLIDE_LAYOUT_RELATIONSHIP,
        target: '../slideLayouts/slideLayout3.xml',
      });
    });
    const duplicateBefore = packageSnapshot(pkg);
    expect(() => resolveSlideLayoutPartUri(
      pkg,
      model.presentationPartUri,
      'BRAND',
      undefined,
    )).toThrow(ModelParseError);
    expect(packageSnapshot(pkg)).toEqual(duplicateBefore);

    const malformedPartUri = '/ppt/slideLayouts/slideLayout-malformed.xml';
    pkg.transaction(() => {
      pkg.setPart(
        malformedPartUri,
        '<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">',
        SLIDE_LAYOUT_CONTENT_TYPE,
      );
      pkg.addRelationship(masterPartUri, {
        type: SLIDE_LAYOUT_RELATIONSHIP,
        target: '../slideLayouts/slideLayout-malformed.xml',
      });
    });
    const malformedBefore = packageSnapshot(pkg);
    expect(() => resolveSlideLayoutPartUri(
      pkg,
      model.presentationPartUri,
      'DEFAULT',
      undefined,
    )).toThrow(ModelParseError);
    expect(packageSnapshot(pkg)).toEqual(malformedBefore);
  });

  it('detects all six OOXML presentation formats from the package content type', async () => {
    const expectedFlags: Record<PresentationFormat, readonly [boolean, boolean, boolean]> = {
      pptx: [false, false, false],
      pptm: [true, false, false],
      ppsx: [false, true, false],
      ppsm: [true, true, false],
      potx: [false, false, true],
      potm: [true, false, true],
    };

    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const model = new PresentationModel(
        await OpcPackage.open(await modelFixture(profile.presentationContentType)),
      );
      expect(model.format).toBe(profile.format);
      expect([
        model.formatProfile.macroEnabled,
        model.formatProfile.slideshow,
        model.formatProfile.template,
      ]).toEqual(expectedFlags[profile.format]);
    }
  });

  it('creates, reads, edits, clears, and no-ops live slide numbers and presentation start', () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    expect(model.firstSlideNumber).toBeUndefined();
    expect(slide.slideNumber).toBeUndefined();

    const options: SlideNumberOptions = {
      align: 'center',
      margin: 0,
      style: { bold: true, color: { kind: 'scheme', value: 'accent1' } },
    };
    slide.slideNumber = options;
    expect(slide.slideNumber).toMatchObject({
      align: 'center',
      margin: { top: 0, right: 0, bottom: 0, left: 0 },
      style: { bold: true, color: { kind: 'scheme', value: 'accent1' } },
    });
    expect(slideNumberCache(pkg, slide.partUri)).toBe('1');

    model.firstSlideNumber = -2;
    expect(model.firstSlideNumber).toBe(-2);
    expect(slideNumberCache(pkg, slide.partUri)).toBe('-2');
    const before = packageSnapshot(pkg);
    model.firstSlideNumber = -2;
    slide.slideNumber = slide.slideNumber;
    expect(packageSnapshot(pkg)).toEqual(before);

    model.firstSlideNumber = undefined;
    expect(model.firstSlideNumber).toBeUndefined();
    expect(slideNumberCache(pkg, slide.partUri)).toBe('1');
    slide.slideNumber = undefined;
    expect(slide.slideNumber).toBeUndefined();
  });

  it('synchronizes direct caches through duplicate, move, delete, and section operations', () => {
    const { pkg, model } = emptyPresentationModel();
    model.firstSlideNumber = 10;
    const first = model.addSlide();
    const second = model.addSlide();
    const third = model.addSlide();
    first.slideNumber = {};
    second.slideNumber = {};
    third.slideNumber = {};
    expect(model.slides.map((slide) => slideNumberCache(pkg, slide.partUri)))
      .toEqual(['10', '11', '12']);

    const duplicate = model.duplicateSlide(0);
    expect(duplicate).not.toBe(first);
    expect(slideNumberCache(pkg, duplicate.partUri)).toBe('13');
    expect(slideNumberCache(pkg, first.partUri)).toBe('10');

    model.moveSlide(model.slides.indexOf(duplicate), 0);
    expect(model.slides[0]).toBe(duplicate);
    expect(model.slides.map((slide) => slideNumberCache(pkg, slide.partUri)))
      .toEqual(['10', '11', '12', '13']);

    const section = model.addSection({ title: 'Numbered', order: 0 });
    model.assignSlideToSection(model.slides.indexOf(second), section.id);
    const beforeSectionRename = model.slides.map((slide) => slideNumberCache(pkg, slide.partUri));
    model.renameSection(section.id, 'Renamed');
    expect(model.slides.map((slide) => slideNumberCache(pkg, slide.partUri)))
      .toEqual(beforeSectionRename);

    model.deleteSlide(model.slides.indexOf(first));
    expect(model.slides.map((slide) => slideNumberCache(pkg, slide.partUri)))
      .toEqual(['10', '11', '12']);
    expect(() => {
      first.slideNumber = { width: 0 };
    }).toThrow(/width/i);
    expect(() => {
      first.slideNumber = {};
    }).toThrow(/current presentation/i);
  });

  it('leaves unsupported direct fields untouched during cache synchronization', () => {
    const { pkg, model } = emptyPresentationModel();
    const supported = model.addSlide();
    const unsupported = model.addSlide();
    supported.slideNumber = {};
    unsupported.slideNumber = {};
    const unsupportedPart = pkg.requirePart(unsupported.partUri);
    const unsupportedXml = new TextDecoder().decode(unsupportedPart.bytes)
      .replace('type="slidenum"', 'type="datetime"');
    pkg.setPart(unsupported.partUri, unsupportedXml, unsupportedPart.contentType);
    const unsupportedBefore = pkg.requirePart(unsupported.partUri).bytes.slice();

    model.firstSlideNumber = 20;
    model.moveSlide(1, 0);
    expect(pkg.requirePart(unsupported.partUri).bytes).toEqual(unsupportedBefore);
    expect(slideNumberCache(pkg, supported.partUri)).toBe('21');
  });

  it('rolls presentation and every changed slide cache back after an injected write failure', () => {
    const { pkg, model } = emptyPresentationModel();
    const first = model.addSlide();
    const second = model.addSlide();
    first.slideNumber = {};
    second.slideNumber = {};
    const before = packageSnapshot(pkg);
    const original = pkg.setPart.bind(pkg);
    let slideWrites = 0;
    const spy = vi.spyOn(pkg, 'setPart').mockImplementation((uri, bytes, contentType) => {
      if (uri.startsWith('/ppt/slides/')) {
        slideWrites += 1;
        if (slideWrites === 2) throw new Error('injected slide-number cache write');
      }
      return original(uri, bytes, contentType);
    });
    expect(() => {
      model.firstSlideNumber = 100;
    }).toThrow('injected slide-number cache write');
    spy.mockRestore();
    expect(packageSnapshot(pkg)).toEqual(before);
    expect(model.firstSlideNumber).toBeUndefined();
    expect(model.slides[0]).toBe(first);
    expect(model.slides[1]).toBe(second);
  });

  it('drops a failed duplicate model after cache synchronization rolls back', () => {
    const { pkg, model } = emptyPresentationModel();
    const source = model.addSlide();
    source.slideNumber = {};
    const before = packageSnapshot(pkg);
    const original = pkg.setPart.bind(pkg);
    let capturedFailedModel: typeof source | undefined;
    const spy = vi.spyOn(pkg, 'setPart').mockImplementation((uri, bytes, contentType) => {
      if (uri === '/ppt/slides/slide2.xml' && pkg.hasPart(uri)) {
        capturedFailedModel = model.slides.at(-1);
        throw new Error('injected duplicate cache write');
      }
      return original(uri, bytes, contentType);
    });
    expect(() => model.duplicateSlide(0)).toThrow('injected duplicate cache write');
    spy.mockRestore();
    expect(packageSnapshot(pkg)).toEqual(before);
    expect(model.slides).toEqual([source]);
    expect(capturedFailedModel).toBeDefined();

    const retry = model.duplicateSlide(0);
    expect(retry).not.toBe(capturedFailedModel);
    expect(model.slides[0]).toBe(source);
    expect(model.slides[1]).toBe(retry);
    expect(slideNumberCache(pkg, retry.partUri)).toBe('2');
  });

  it('stores strict frozen slide default text colors without package mutation', () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const input = { kind: 'srgb' as const, value: '#ff3399' };
    const before = packageSnapshot(pkg);

    slide.color = input;
    expect(slide.color).toEqual({ kind: 'srgb', value: 'FF3399' });
    expect(slide.color).not.toBe(input);
    expect(Object.isFrozen(slide.color)).toBe(true);
    expect(packageSnapshot(pkg)).toEqual(before);

    const snapshot = slide.color;
    slide.color = { kind: 'srgb', value: 'FF3399' };
    expect(slide.color).toBe(snapshot);
    expect(packageSnapshot(pkg)).toEqual(before);

    slide.color = { kind: 'scheme', value: 'accent2' };
    expect(slide.color).toEqual({ kind: 'scheme', value: 'accent2' });
    expect(Object.isFrozen(slide.color)).toBe(true);
    slide.color = undefined;
    expect(slide.color).toBeUndefined();
    slide.color = undefined;
    expect(packageSnapshot(pkg)).toEqual(before);
  });

  it('rejects unsafe slide default text colors without changing prior state', () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    slide.color = { kind: 'scheme', value: 'accent1' };
    const snapshot = slide.color;
    const before = packageSnapshot(pkg);
    const inherited = Object.create({ kind: 'srgb', value: 'FF3399' }) as object;
    const accessor = Object.defineProperty({ kind: 'srgb' }, 'value', {
      enumerable: true,
      get: () => 'FF3399',
    });
    const nullPrototype = Object.assign(Object.create(null) as object, {
      kind: 'srgb',
      value: '#112233',
    });
    slide.color = nullPrototype as never;
    expect(slide.color).toEqual({ kind: 'srgb', value: '112233' });
    slide.color = snapshot;
    const restored = slide.color;

    const invalid = [
      null,
      'FF3399',
      [],
      { kind: 'srgb' },
      { kind: 'srgb', value: 'FFF' },
      { kind: 'scheme', value: 'unknown' },
      { kind: 'srgb', value: 'FF3399', extra: true },
      inherited,
      accessor,
      { kind: 'srgb', value: 'FF3399', [Symbol('extra')]: true },
    ];
    for (const candidate of invalid) {
      expect(() => {
        slide.color = candidate as never;
      }).toThrow();
      expect(slide.color).toBe(restored);
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    const { proxy, revoke } = Proxy.revocable({ kind: 'srgb', value: 'FF3399' }, {});
    revoke();
    expect(() => {
      slide.color = proxy as never;
    }).toThrow();
    expect(slide.color).toBe(restored);
    expect(packageSnapshot(pkg)).toEqual(before);
  });

  it('copies, retains, and clears slide default text colors through lifecycle operations', () => {
    const { model } = emptyPresentationModel();
    const source = model.addSlide();
    const sibling = model.addSlide();
    source.color = { kind: 'scheme', value: 'accent2' };
    expect(sibling.color).toBeUndefined();

    const duplicate = model.duplicateSlide(model.slides.indexOf(source));
    expect(duplicate.color).toEqual(source.color);
    expect(duplicate.color).toBe(source.color);
    model.moveSlide(model.slides.indexOf(duplicate), 0);
    expect(model.slides[0]).toBe(duplicate);
    expect(duplicate.color).toEqual({ kind: 'scheme', value: 'accent2' });

    const sourcePartUri = source.partUri;
    model.deleteSlide(model.slides.indexOf(source));
    expect(source.color).toBeUndefined();
    const replacement = model.addSlide();
    expect(replacement.partUri).toBe(sourcePartUri);
    expect(replacement.color).toBeUndefined();
    expect(duplicate.color).toEqual({ kind: 'scheme', value: 'accent2' });
  });

  it('commits transient color lifecycle only after package transactions succeed', () => {
    const { pkg, model } = emptyPresentationModel();
    const source = model.addSlide();
    source.color = { kind: 'srgb', value: 'FF3399' };
    const originalSetPart = pkg.setPart.bind(pkg);
    const duplicateFailure = vi.spyOn(pkg, 'setPart').mockImplementation(
      (uri, bytes, contentType) => {
        if (uri === '/ppt/presentation.xml' && pkg.hasPart('/ppt/slides/slide2.xml')) {
          throw new Error('injected transient duplicate failure');
        }
        return originalSetPart(uri, bytes, contentType);
      },
    );
    expect(() => model.duplicateSlide(0)).toThrow('injected transient duplicate failure');
    duplicateFailure.mockRestore();
    expect(model.slides).toEqual([source]);
    expect(source.color).toEqual({ kind: 'srgb', value: 'FF3399' });

    const duplicate = model.duplicateSlide(0);
    expect(duplicate.color).toBe(source.color);
    const originalDeletePart = pkg.deletePart.bind(pkg);
    const deleteFailure = vi.spyOn(pkg, 'deletePart').mockImplementation((uri) => {
      if (uri === source.partUri) throw new Error('injected transient delete failure');
      return originalDeletePart(uri);
    });
    expect(() => model.deleteSlide(model.slides.indexOf(source)))
      .toThrow('injected transient delete failure');
    deleteFailure.mockRestore();
    expect(model.slides.includes(source)).toBe(true);
    expect(source.color).toEqual({ kind: 'srgb', value: 'FF3399' });

    model.deleteSlide(model.slides.indexOf(source));
    expect(source.color).toBeUndefined();
    expect(duplicate.color).toEqual({ kind: 'srgb', value: 'FF3399' });
  });

  it('materializes slide default colors into only newly created plain and rich text', () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    slide.color = { kind: 'srgb', value: '#ff3399' };
    const plain = slide.addText('First\nSecond');
    expect(plain.richText.map(({ runs }) => runs[0]?.style?.color)).toEqual([
      { kind: 'srgb', value: 'FF3399' },
      { kind: 'srgb', value: 'FF3399' },
    ]);
    const plainXml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(plainXml.match(/<a:srgbClr val="FF3399"\/>/g)).toHaveLength(2);

    const beforeColorChange = pkg.requirePart(slide.partUri).bytes;
    slide.color = { kind: 'scheme', value: 'accent1' };
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeColorChange);
    const rich = slide.addRichText([{
      runs: [
        { text: 'Inherited' },
        { text: 'Override', style: { color: { kind: 'srgb', value: '00AA00' } } },
        { text: 'Transparent inherited', style: { transparency: 25 } },
      ],
    }]);
    expect(rich.richText[0]?.runs.map(({ style }) => style?.color)).toEqual([
      { kind: 'scheme', value: 'accent1' },
      { kind: 'srgb', value: '00AA00' },
      { kind: 'scheme', value: 'accent1' },
    ]);
    expect(rich.richText[0]?.runs.map(({ style }) => style?.transparency)).toEqual([
      undefined,
      undefined,
      25,
    ]);
    const richXml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(richXml).toContain(
      '<a:schemeClr val="accent1"><a:alpha val="75000"/></a:schemeClr>',
    );

    slide.color = undefined;
    const cleared = slide.addText('Theme default');
    expect(cleared.richText[0]?.runs[0]?.style?.color).toEqual({
      kind: 'scheme',
      value: 'tx1',
    });
    expect(plain.richText.map(({ runs }) => runs[0]?.style?.color)).toEqual([
      { kind: 'srgb', value: 'FF3399' },
      { kind: 'srgb', value: 'FF3399' },
    ]);
  });

  it('does not apply slide defaults to existing-shape edits or tables', () => {
    const { model } = emptyPresentationModel();
    const slide = model.addSlide();
    const existing = slide.addText('Existing');
    slide.color = { kind: 'scheme', value: 'accent6' };
    existing.richText = [{ runs: [{ text: 'Edited existing' }] }];
    expect(existing.richText[0]?.runs[0]?.style?.color).toEqual({
      kind: 'scheme',
      value: 'tx1',
    });

    const table = slide.addTable([['Table default']]);
    const { xml, element } = slide.resolveShape(table.id);
    const tableXml = xml.original(element);
    expect(tableXml).toContain('<a:schemeClr val="tx1"/>');
    expect(tableXml).not.toContain('<a:schemeClr val="accent6"/>');
    expect(slide.color).toEqual({ kind: 'scheme', value: 'accent6' });
  });

  it('creates, reads, edits, clears, no-ops, and rolls back live slide backgrounds', () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    expect(slide.background).toBeUndefined();

    slide.background = { kind: 'none' };
    expect(slide.background).toEqual({ kind: 'none' });
    slide.background = {
      kind: 'solid',
      color: { kind: 'srgb', value: 'ff3399' },
      transparency: 50,
    };
    expect(slide.background).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF3399' },
      transparency: 50,
    });

    const beforeNoOp = pkg.requirePart(slide.partUri).bytes;
    const beforeNoOpJournal = [...pkg.mutations];
    slide.background = {
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF3399' },
      transparency: 50,
    };
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeNoOp);
    expect(pkg.mutations).toEqual(beforeNoOpJournal);

    expect(() => pkg.transaction(() => {
      slide.background = {
        kind: 'path-gradient',
        path: 'circle',
        stops: [
          { offset: 0, color: 'FFFFFF' },
          { offset: 1, color: '000000' },
        ],
      };
      throw new Error('rollback background');
    })).toThrow('rollback background');
    expect(slide.background).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF3399' },
      transparency: 50,
    });

    slide.background = undefined;
    expect(slide.background).toBeUndefined();
  });

  it('isolates duplicated image backgrounds and garbage-collects only the final target', () => {
    const { pkg, model } = emptyPresentationModel();
    const source = model.addSlide();
    source.background = {
      kind: 'image',
      contentType: 'image/png',
      bytes: Uint8Array.of(1, 2, 3),
    };
    const sourceRelationship = source.relationships.find(({ type }) =>
      type === IMAGE_RELATIONSHIP)!;
    const sharedTarget = sourceRelationship.resolvedTarget!;

    const duplicate = model.duplicateSlide(0);
    expect(duplicate.relationships.find(({ type }) => type === IMAGE_RELATIONSHIP)?.resolvedTarget)
      .toBe(sharedTarget);
    duplicate.background = {
      kind: 'image',
      contentType: 'image/png',
      bytes: Uint8Array.of(9, 8, 7),
    };
    const duplicateTarget = duplicate.relationships.find(({ type }) =>
      type === IMAGE_RELATIONSHIP)!.resolvedTarget!;
    expect(duplicateTarget).not.toBe(sharedTarget);
    expect(source.background).toMatchObject({ bytes: Uint8Array.of(1, 2, 3) });
    expect(duplicate.background).toMatchObject({ bytes: Uint8Array.of(9, 8, 7) });

    model.deleteSlide(1);
    expect(pkg.hasPart(duplicateTarget)).toBe(false);
    expect(pkg.hasPart(sharedTarget)).toBe(true);
    model.deleteSlide(0);
    expect(pkg.hasPart(sharedTarget)).toBe(false);
  });

  it.each([
    ['source then duplicate', 0],
    ['duplicate then source', 1],
  ] as const)('retains a shared image background when deleting %s', (_name, firstIndex) => {
    const { pkg, model } = emptyPresentationModel();
    const source = model.addSlide();
    source.background = {
      kind: 'image',
      contentType: 'image/png',
      bytes: Uint8Array.of(1, 2, 3),
    };
    const target = source.relationships.find(({ type }) =>
      type === IMAGE_RELATIONSHIP)!.resolvedTarget!;
    model.duplicateSlide(0);

    model.deleteSlide(firstIndex);
    expect(pkg.hasPart(target)).toBe(true);
    model.deleteSlide(0);
    expect(pkg.hasPart(target)).toBe(false);
  });

  it('edits a PptxGenJS-style image background without rewriting unrelated slide XML', () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    slide.addShape('rect', { name: 'Keep neighbor' });
    slide.background = {
      kind: 'image',
      contentType: 'image/png',
      bytes: Uint8Array.of(1, 2, 3),
    };
    const relationship = slide.relationships.find(({ type }) =>
      type === IMAGE_RELATIONSHIP)!;
    const target = relationship.resolvedTarget!;
    const part = pkg.requirePart(slide.partUri);
    const source = new TextDecoder().decode(part.bytes);
    const canonical = `<a:blipFill><a:blip r:embed="${relationship.id}"/>`
      + '<a:stretch><a:fillRect/></a:stretch></a:blipFill>';
    const pptxGenJS = `<a:blipFill dpi="0" rotWithShape="1">`
      + `<a:blip r:embed="${relationship.id}"><a:lum/></a:blip>`
      + '<a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>';
    expect(source).toContain(canonical);
    pkg.setPart(
      slide.partUri,
      source.replace(canonical, pptxGenJS).replace(
        '</p:sld>',
        '<p:extLst><p:ext uri="urn:keep"><x:keep xmlns:x="urn:keep">UNCHANGED</x:keep>'
          + '</p:ext></p:extLst></p:sld>',
      ),
      part.contentType,
    );

    expect(slide.background).toEqual({
      kind: 'image',
      contentType: 'image/png',
      bytes: Uint8Array.of(1, 2, 3),
    });
    slide.background = {
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF3399' },
      transparency: 50,
    };

    expect(pkg.hasPart(target)).toBe(false);
    expect(slide.relationships.some(({ id }) => id === relationship.id)).toBe(false);
    const updated = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(updated).toContain('name="Keep neighbor"');
    expect(updated).toContain('<a:effectLst/>');
    expect(updated).toContain('<x:keep xmlns:x="urn:keep">UNCHANGED</x:keep>');
    expect(updated).not.toContain('dpi="0"');
  });

  it('creates preset shapes through the live model and rejects unsafe shape ids atomically', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const rectangle = slide.addShape('rect');
    const star = slide.addShape('star5', {
      name: 'Model star',
      x: inches(2),
      width: inches(3),
    });
    expect([rectangle.id, star.id]).toEqual([2, 3]);
    expect(slide.shapes).toEqual([rectangle, star]);
    expect(PRESET_SHAPE_TYPES).toContain('star5');

    const part = pkg.requirePart(slide.partUri);
    const malformed = new TextDecoder().decode(part.bytes).replace('id="1"', 'id="not-an-id"');
    pkg.setPart(slide.partUri, malformed, part.contentType);
    const before = pkg.requirePart(slide.partUri).bytes.slice();
    const journal = [...pkg.mutations];
    expect(() => slide.addShape('ellipse')).toThrow(ModelParseError);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(before);
    expect(pkg.mutations).toEqual(journal);
  });

  it('creates embedded raster images from a zero-input package with immediate live state', () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const source = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
    const image = slide.addImage(source, {
      contentType: 'image/png',
      name: 'Revenue & <logo>',
      altText: 'Quarterly & annual',
      x: inches(1),
      y: inches(2),
      width: inches(3),
      height: inches(2),
      rotation: degrees(45),
      flipHorizontal: true,
      sourceRectangle: {
        left: 25,
        top: -10,
        right: 5,
        bottom: 0,
      },
    });

    expect(image).toBeInstanceOf(ImageModel);
    expect(image.kind).toBe('image');
    expect(image.sourceRectangle).toEqual({
      left: 25,
      top: -10,
      right: 5,
      bottom: 0,
    });
    expect(slide.shapes).toEqual([image]);
    expect(slide.shapes[0]).toBe(image);
    expect(image.name).toBe('Revenue & <logo>');
    expect(image.transform).toEqual({
      x: 914_400,
      y: 1_828_800,
      width: 2_743_200,
      height: 1_828_800,
      rotation: 2_700_000,
      flipHorizontal: true,
      flipVertical: false,
    });
    expect(image.sourcePartUri).toBe('/ppt/media/image1.png');
    const relationship = slide.relationships.find(({ type }) => type === IMAGE_RELATIONSHIP);
    expect(relationship).toMatchObject({
      type: IMAGE_RELATIONSHIP,
      target: '../media/image1.png',
      targetMode: 'Internal',
      resolvedTarget: '/ppt/media/image1.png',
    });

    const slideSource = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(slideSource).toContain(
      '<p:cNvPr id="2" name="Revenue &amp; &lt;logo&gt;" ' +
      'descr="Quarterly &amp; annual"/>',
    );
    expect(slideSource).toContain(`<a:blip r:embed="${relationship!.id}"/>`);
    expect(slideSource).toContain(
      '<a:srcRect l="25000" t="-10000" r="5000" b="0"/>',
    );
    expect(slideSource.indexOf('<a:blip ')).toBeLessThan(slideSource.indexOf('<a:srcRect '));
    expect(slideSource.indexOf('<a:srcRect ')).toBeLessThan(
      slideSource.indexOf('<a:stretch>'),
    );
    expect(slideSource).toContain('<a:xfrm rot="2700000" flipH="1">');
    expect(slideSource).toContain('<a:picLocks noChangeAspect="1"/>');
    expect(slideSource).toContain('<a:stretch><a:fillRect/></a:stretch>');

    source.fill(0);
    const part = pkg.requirePart('/ppt/media/image1.png');
    expect(part.contentType).toBe('image/png');
    expect(part.bytes).toEqual(new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]));
  });

  it('creates stable live media shapes and reconciles semantic kind changes', async () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const audio = await slide.addAudio(Uint8Array.of(1, 2, 3), {
      name: 'Stable narration',
      altText: 'Spoken overview',
      contentType: 'audio/mpeg',
      poster: Uint8Array.of(4, 5, 6),
      posterContentType: 'image/png',
      play: 'auto',
      loop: true,
      volume: 0.25,
    });
    const beforeRead = packageSnapshot(pkg);

    expect(audio).toBeInstanceOf(MediaModel);
    expect(audio.kind).toBe('audio');
    expect(audio.shapeId).toBe(audio.id);
    expect(audio.slidePartUri).toBe(slide.partUri);
    expect(audio.name).toBe('Stable narration');
    expect(audio.altText).toBe('Spoken overview');
    expect(audio.mediaPartUri).toMatch(/\.mp3$/);
    expect(audio.posterPartUri).toMatch(/\.png$/);
    expect(audio.settings).toEqual({
      play: 'auto',
      loop: true,
      hideWhenStopped: false,
      volume: 0.25,
    });
    expect(Object.isFrozen(audio.settings)).toBe(true);
    expect(slide.media[0]).toBe(audio);
    expect(slide.shapes[0]).toBe(audio);
    expect(model.slides[0]!.media[0]).toBe(audio);
    expect(packageSnapshot(pkg)).toEqual(beforeRead);

    const duplicate = model.duplicateSlide(0);
    const duplicateAudio = duplicate.media[0]!;
    expect(duplicateAudio).toBeInstanceOf(MediaModel);
    expect(duplicateAudio).not.toBe(audio);
    expect(duplicateAudio.mediaPartUri).toBe(audio.mediaPartUri);
    expect(duplicateAudio.posterPartUri).toBe(audio.posterPartUri);
    model.moveSlide(1, 0);
    expect(model.slides[0]).toBe(duplicate);
    expect(model.slides[0]!.media[0]).toBe(duplicateAudio);
    expect(model.slides[1]).toBe(slide);
    expect(model.slides[1]!.media[0]).toBe(audio);

    expect(() => pkg.transaction(() => {
      const part = pkg.requirePart(slide.partUri);
      pkg.setPart(
        slide.partUri,
        new TextDecoder().decode(part.bytes).replace('Stable narration', 'Rolled back'),
        part.contentType,
      );
      expect(slide.media[0]).toBe(audio);
      throw new Error('rollback identity');
    })).toThrow('rollback identity');
    expect(slide.media[0]).toBe(audio);
    expect(audio.name).toBe('Stable narration');

    const kindRelationship = slide.relationships.find(({ type }) => type.endsWith('/audio'))!;
    const mediaUri = audio.mediaPartUri!;
    audio.settings = undefined;
    const slidePart = pkg.requirePart(slide.partUri);
    pkg.transaction(() => {
      pkg.setPart(
        slide.partUri,
        new TextDecoder().decode(slidePart.bytes)
          .replace('<a:audioFile ', '<a:videoFile '),
        slidePart.contentType,
      );
      pkg.setPart(mediaUri, pkg.requirePart(mediaUri).bytes, 'video/mp4');
      pkg.updateRelationship(slide.partUri, kindRelationship.id, {
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video',
      });
    });
    const video = slide.media[0]!;
    expect(video).toBeInstanceOf(MediaModel);
    expect(video.kind).toBe('video');
    expect(video).not.toBe(audio);
    expect(() => audio.mediaPartUri).toThrow(/changed semantic kind/);

    new MediaCodec(pkg).delete(slide.partUri, video.shapeId);
    expect(slide.media).toEqual([]);
    expect(() => video.name).toThrow(/was not found/);
  });

  it('edits live media metadata, transform, and playback preferences atomically', async () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const media = await slide.addAudio(Uint8Array.of(1, 2, 3), {
      name: 'Original',
      altText: 'Original description',
      contentType: 'audio/mpeg',
      play: 'click',
    });
    const createdSource = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(createdSource).toContain('nativeVersion="1"');
    expect(createdSource).toContain('cmd="playFrom(0.0)"');

    const noOp = packageSnapshot(pkg);
    media.name = 'Original';
    media.altText = 'Original description';
    media.settings = { play: 'click', loop: false, hideWhenStopped: false, volume: 1 };
    expect(packageSnapshot(pkg)).toEqual(noOp);

    media.name = 'Audio & <narration>';
    media.altText = '';
    media.setTransform({
      x: inches(1),
      y: inches(2),
      width: inches(3),
      height: inches(4),
      rotation: degrees(90),
      flipHorizontal: true,
      flipVertical: true,
    });
    media.settings = {
      play: 'auto',
      loop: true,
      hideWhenStopped: true,
      volume: 0.125,
    };
    expect(media.name).toBe('Audio & <narration>');
    expect(media.altText).toBe('');
    expect(media.transform).toEqual({
      x: inches(1),
      y: inches(2),
      width: inches(3),
      height: inches(4),
      rotation: degrees(90),
      flipHorizontal: true,
      flipVertical: true,
    });
    expect(media.settings).toEqual({
      play: 'auto',
      loop: true,
      hideWhenStopped: true,
      volume: 0.125,
    });
    expect(slide.media[0]).toBe(media);
    const editedSource = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(editedSource).toContain('name="Audio &amp; &lt;narration&gt;" descr=""');
    expect(editedSource).toContain('repeatCount="indefinite"');
    expect(editedSource).toContain('showWhenStopped="0"');
    expect(editedSource).toContain('vol="12500"');

    media.altText = undefined;
    media.settings = undefined;
    expect(media.altText).toBeUndefined();
    expect(media.settings).toEqual({});
    const clearedSource = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(clearedSource).not.toContain(' descr=');
    expect(clearedSource).not.toContain('px:playback');
    expect(clearedSource).not.toContain('<p:timing>');

    const beforeInvalid = packageSnapshot(pkg);
    let reads = 0;
    const unsafe = Object.defineProperty({}, 'play', {
      enumerable: true,
      get() { reads += 1; return 'auto'; },
    });
    expect(() => { media.settings = unsafe as never; }).toThrow(/data property/);
    expect(() => { media.name = '\u0001'; }).toThrow(/invalid XML/);
    expect(reads).toBe(0);
    expect(packageSnapshot(pkg)).toEqual(beforeInvalid);

    expect(() => pkg.transaction(() => {
      media.name = 'Rolled back';
      media.altText = 'Rolled back';
      media.settings = { play: 'auto', loop: true };
      throw new Error('rollback media editing');
    })).toThrow('rollback media editing');
    expect(media.name).toBe('Audio & <narration>');
    expect(media.altText).toBeUndefined();
    expect(media.settings).toEqual({});

    const duplicate = model.duplicateSlide(0);
    const duplicateMedia = duplicate.media[0]!;
    duplicateMedia.name = 'Duplicate only';
    duplicateMedia.altText = 'Duplicate description';
    duplicateMedia.settings = { play: 'auto', volume: 0.5 };
    expect(duplicateMedia.name).toBe('Duplicate only');
    expect(media.name).toBe('Audio & <narration>');
    expect(media.altText).toBeUndefined();
    expect(media.settings).toEqual({});

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedMedia = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!.media[0]!;
    expect(reopenedMedia.name).toBe('Audio & <narration>');
    expect(reopenedMedia.transform).toEqual(media.transform);
    expect(reopenedMedia.settings).toEqual({});
  });

  it('isolates native media timing through duplicate, move, replacement, and removal', async () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const media = await slide.addVideo(Uint8Array.of(1, 2, 3), {
      contentType: 'video/mp4',
      poster: Uint8Array.of(4, 5),
      posterContentType: 'image/png',
      play: 'auto',
      loop: true,
      volume: 0.5,
    });
    const timing = (partUri: string): string => {
      const xml = LosslessXmlDocument.parse(pkg.requirePart(partUri).bytes);
      return xml.original(xml.elements('timing')[0]!);
    };
    const sourceTiming = timing(slide.partUri);
    const sourceMedia = media.mediaPartUri!;
    const sourcePoster = media.posterPartUri!;
    const sourceMediaBytes = pkg.requirePart(sourceMedia).bytes.slice();
    const duplicate = model.duplicateSlide(0);
    const duplicateMedia = duplicate.media[0]!;

    expect(timing(duplicate.partUri)).toBe(sourceTiming);
    model.moveSlide(1, 0);
    expect(model.slides[0]).toBe(duplicate);
    expect(timing(duplicate.partUri)).toBe(sourceTiming);
    expect(timing(slide.partUri)).toBe(sourceTiming);

    duplicateMedia.settings = {
      play: 'click',
      loop: false,
      hideWhenStopped: true,
      volume: 0.25,
    };
    expect(timing(duplicate.partUri)).not.toBe(sourceTiming);
    expect(timing(slide.partUri)).toBe(sourceTiming);
    await duplicateMedia.replaceSource(Uint8Array.of(6, 7), { contentType: 'video/mp4' });
    await duplicateMedia.replacePoster(Uint8Array.of(8, 9), { contentType: 'image/gif' });
    expect(timing(slide.partUri)).toBe(sourceTiming);
    expect(pkg.requirePart(sourceMedia).bytes).toEqual(sourceMediaBytes);
    expect(media.posterPartUri).toBe(sourcePoster);

    duplicateMedia.remove();
    expect(timing(slide.partUri)).toBe(sourceTiming);
    expect(pkg.hasPart(sourceMedia)).toBe(true);
    media.remove();
    expect(new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes)).not.toContain('<p:timing>');
  });

  it('replaces live media sources across embedded and external modes with COW isolation', async () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const media = await slide.addAudio(Uint8Array.of(1, 2, 3), {
      name: 'Replace me',
      altText: 'Keep metadata',
      contentType: 'audio/mpeg',
      poster: Uint8Array.of(9),
      posterContentType: 'image/png',
      play: 'auto',
      x: inches(1),
      y: inches(2),
      width: inches(3),
      height: inches(1),
    });
    const identity = media;
    const originalUri = media.mediaPartUri!;
    const posterUri = media.posterPartUri;
    const appearance = {
      name: media.name,
      altText: media.altText,
      transform: media.transform,
      settings: media.settings,
    };

    expect(await media.replaceSource(Uint8Array.of(4, 5, 6), {
      contentType: 'audio/mpeg',
    })).toBe(media);
    expect(media.mediaPartUri).toBe(originalUri);
    expect(pkg.requirePart(originalUri).bytes).toEqual(Uint8Array.of(4, 5, 6));

    await media.replaceSource('https://example.com/narration.mp3');
    expect(media).toBe(identity);
    expect(media.mediaPartUri).toBeUndefined();
    expect(media.externalUrl).toBe('https://example.com/narration.mp3');
    expect(pkg.hasPart(originalUri)).toBe(false);
    expect(media.posterPartUri).toBe(posterUri);
    expect({
      name: media.name,
      altText: media.altText,
      transform: media.transform,
      settings: media.settings,
    }).toEqual(appearance);

    await media.replaceSource('data:audio/wav;base64,BwgJ', { fileName: 'voice.wav' });
    expect(media.externalUrl).toBeUndefined();
    expect(media.mediaPartUri).toMatch(/\.wav$/);
    expect(pkg.requirePart(media.mediaPartUri!)).toMatchObject({
      contentType: 'audio/wav',
      bytes: Uint8Array.of(7, 8, 9),
    });
    const sharedWav = media.mediaPartUri!;
    const duplicate = model.duplicateSlide(0);
    const duplicateMedia = duplicate.media[0]!;
    expect(duplicateMedia.mediaPartUri).toBe(sharedWav);
    await duplicateMedia.replaceSource(Uint8Array.of(10, 11), {
      contentType: 'audio/wav',
      fileName: 'changed.wav',
    });
    expect(duplicateMedia.mediaPartUri).not.toBe(sharedWav);
    expect(pkg.requirePart(sharedWav).bytes).toEqual(Uint8Array.of(7, 8, 9));
    expect(pkg.requirePart(duplicateMedia.mediaPartUri!).bytes).toEqual(Uint8Array.of(10, 11));
    expect(media.mediaPartUri).toBe(sharedWav);

    const dedup = await slide.addAudio(Uint8Array.of(12, 13), {
      contentType: 'audio/wav',
      fileName: 'dedup.wav',
    });
    await duplicateMedia.replaceSource(Uint8Array.of(12, 13), {
      contentType: 'audio/wav',
      fileName: 'dedup.wav',
    });
    expect(duplicateMedia.mediaPartUri).toBe(dedup.mediaPartUri);

    const beforeInvalid = packageSnapshot(pkg);
    await expect(duplicateMedia.replaceSource('data:video/mp4;base64,AQ=='))
      .rejects.toThrow(/Unsupported audio content type/);
    await expect(duplicateMedia.replaceSource(Uint8Array.of(1), { x: 1 } as never))
      .rejects.toThrow(/unsupported property/);
    expect(packageSnapshot(pkg)).toEqual(beforeInvalid);
    expect(duplicate.media[0]).toBe(duplicateMedia);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedDuplicate = reopened.slides.find(({ partUri }) => partUri === duplicate.partUri)!.media[0]!;
    expect(reopenedDuplicate.mediaPartUri).toBe(dedup.mediaPartUri);
    expect(reopenedDuplicate.name).toBe('Replace me');
    expect(reopenedDuplicate.posterPartUri).toBe(posterUri);
  });

  it('isolates shared media relationship ids and canonicalizes legacy primary roles', async () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const first = await slide.addAudio(Uint8Array.of(1, 2), {
      name: 'Shared first',
      contentType: 'audio/mpeg',
      poster: Uint8Array.of(3),
      posterContentType: 'image/png',
    });
    const originalTarget = first.mediaPartUri!;
    const part = pkg.requirePart(slide.partUri);
    const source = new TextDecoder().decode(part.bytes);
    const picture = source.match(/<p:pic>.*?<\/p:pic>/)?.[0];
    expect(picture).toBeTruthy();
    pkg.setPart(
      slide.partUri,
      source.replace(
        '</p:spTree>',
        `${picture!.replace('id="2"', 'id="3"').replace('name="Shared first"', 'name="Shared second"')}</p:spTree>`,
      ),
      part.contentType,
    );
    const second = slide.media.find(({ id }) => id === 3)!;
    expect(second.mediaPartUri).toBe(originalTarget);
    await second.replaceSource(Uint8Array.of(4, 5), { contentType: 'audio/mpeg' });
    expect(first.mediaPartUri).toBe(originalTarget);
    expect(second.mediaPartUri).not.toBe(originalTarget);
    expect(pkg.requirePart(originalTarget).bytes).toEqual(Uint8Array.of(1, 2));
    expect(pkg.requirePart(second.mediaPartUri!).bytes).toEqual(Uint8Array.of(4, 5));
    expect(slide.relationships.filter(({ type }) => type.endsWith('/audio'))).toHaveLength(2);
    expect(slide.relationships.filter(
      ({ type }) => type === 'http://schemas.microsoft.com/office/2007/relationships/media',
    )).toHaveLength(2);
    const firstPoster = first.posterPartUri!;
    expect(second.posterPartUri).toBe(firstPoster);
    await second.replacePoster(Uint8Array.of(6), {
      contentType: 'image/gif',
      fileName: 'shared.gif',
    });
    expect(first.posterPartUri).toBe(firstPoster);
    expect(second.posterPartUri).not.toBe(firstPoster);
    expect(pkg.requirePart(firstPoster).bytes).toEqual(Uint8Array.of(3));
    expect(pkg.requirePart(second.posterPartUri!).bytes).toEqual(Uint8Array.of(6));
    expect(slide.relationships.filter(({ type }) => type.endsWith('/image'))).toHaveLength(2);

    const kindRelationship = slide.relationships.find(
      ({ id, type }) => id !== undefined && type.endsWith('/audio'),
    )!;
    const firstPart = pkg.requirePart(first.mediaPartUri!);
    pkg.transaction(() => {
      const slideBytes = pkg.requirePart(slide.partUri);
      pkg.setPart(
        slide.partUri,
        new TextDecoder().decode(slideBytes.bytes).replace('<a:audioFile ', '<a:videoFile '),
        slideBytes.contentType,
      );
      pkg.updateRelationship(slide.partUri, kindRelationship.id, {
        type: 'http://schemas.microsoft.com/office/2007/relationships/media',
      });
      pkg.setPart(firstPart.uri, firstPart.bytes, 'audio/mp3');
    });
    expect(first.kind).toBe('audio');
    await first.replaceSource(Uint8Array.of(6, 7), { contentType: 'audio/mpeg' });
    const canonical = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(canonical).toContain('<a:audioFile ');
    expect(slide.relationships.some(({ type, resolvedTarget }) =>
      type.endsWith('/audio') && resolvedTarget === first.mediaPartUri)).toBe(true);
    expect(slide.relationships.some(({ type, resolvedTarget }) =>
      type === 'http://schemas.microsoft.com/office/2007/relationships/media'
      && resolvedTarget === first.mediaPartUri)).toBe(true);
    expect(second.mediaPartUri).not.toBe(first.mediaPartUri);
  });

  it('replaces and resets live media posters with COW and dedup isolation', async () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const media = await slide.addVideo(Uint8Array.of(1, 2), {
      name: 'Poster lifecycle',
      altText: 'Keep poster metadata',
      contentType: 'video/mp4',
      poster: Uint8Array.of(3, 4),
      posterContentType: 'image/png',
      play: 'auto',
    });
    const mediaUri = media.mediaPartUri;
    const identity = media;
    const originalPoster = media.posterPartUri!;
    await media.replacePoster(Uint8Array.of(4, 5), {
      contentType: 'image/png',
      fileName: 'cover.png',
    });
    expect(media.posterPartUri).toBe(originalPoster);
    expect(pkg.requirePart(originalPoster).bytes).toEqual(Uint8Array.of(4, 5));
    expect(await media.replacePoster('data:image/jpeg;base64,BQYH', {
      fileName: 'cover.jpeg',
    })).toBe(media);
    expect(media).toBe(identity);
    expect(media.posterPartUri).toMatch(/\.jpeg$/);
    expect(media.posterPartUri).not.toBe(originalPoster);
    expect(pkg.hasPart(originalPoster)).toBe(false);
    expect(pkg.requirePart(media.posterPartUri!)).toMatchObject({
      contentType: 'image/jpeg',
      bytes: Uint8Array.of(5, 6, 7),
    });
    expect(media.mediaPartUri).toBe(mediaUri);
    expect(media.name).toBe('Poster lifecycle');
    expect(media.altText).toBe('Keep poster metadata');
    expect(media.settings.play).toBe('auto');

    await media.replacePoster();
    expect(media.posterPartUri).toMatch(/\.png$/);
    expect(pkg.requirePart(media.posterPartUri!).bytes.slice(0, 8))
      .toEqual(Uint8Array.of(137, 80, 78, 71, 13, 10, 26, 10));
    const sharedDefault = media.posterPartUri!;
    const beforeRepeatedReset = packageSnapshot(pkg);
    await media.replacePoster();
    expect(media.posterPartUri).toBe(sharedDefault);
    expect(packageSnapshot(pkg)).toEqual(beforeRepeatedReset);
    const duplicate = model.duplicateSlide(0);
    const duplicateMedia = duplicate.media[0]!;
    expect(duplicateMedia.posterPartUri).toBe(sharedDefault);
    await duplicateMedia.replacePoster(Uint8Array.of(8, 9), {
      contentType: 'image/gif',
      fileName: 'cover.gif',
    });
    expect(duplicateMedia.posterPartUri).not.toBe(sharedDefault);
    expect(media.posterPartUri).toBe(sharedDefault);
    expect(pkg.requirePart(sharedDefault).bytes.slice(0, 8))
      .toEqual(Uint8Array.of(137, 80, 78, 71, 13, 10, 26, 10));

    const dedup = await slide.addVideo(Uint8Array.of(10), {
      contentType: 'video/mp4',
      poster: Uint8Array.of(11, 12),
      posterContentType: 'image/gif',
    });
    await duplicateMedia.replacePoster(Uint8Array.of(11, 12), {
      contentType: 'image/gif',
    });
    expect(duplicateMedia.posterPartUri).toBe(dedup.posterPartUri);

    const beforeInvalid = packageSnapshot(pkg);
    await expect(duplicateMedia.replacePoster('https://example.com/poster.png'))
      .rejects.toThrow(/External poster URLs/);
    await expect(duplicateMedia.replacePoster(Uint8Array.of(1), { contentType: 'video/mp4' }))
      .rejects.toThrow(/Unsupported poster content type/);
    expect(packageSnapshot(pkg)).toEqual(beforeInvalid);
    expect(duplicate.media[0]).toBe(duplicateMedia);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedDuplicate = reopened.slides.find(({ partUri }) => partUri === duplicate.partUri)!.media[0]!;
    expect(reopenedDuplicate.posterPartUri).toBe(dedup.posterPartUri);
    expect(reopenedDuplicate.mediaPartUri).toBe(mediaUri);
  });

  it('removes live media without breaking shared relationships or payloads', async () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const first = await slide.addAudio(Uint8Array.of(1, 2), {
      name: 'Shared removal',
      contentType: 'audio/mpeg',
      poster: Uint8Array.of(3),
      posterContentType: 'image/png',
    });
    const mediaUri = first.mediaPartUri!;
    const posterUri = first.posterPartUri!;
    const slidePart = pkg.requirePart(slide.partUri);
    const source = new TextDecoder().decode(slidePart.bytes);
    const picture = source.match(/<p:pic>.*?<\/p:pic>/)?.[0];
    expect(picture).toBeTruthy();
    pkg.setPart(
      slide.partUri,
      source.replace(
        '</p:spTree>',
        `${picture!.replace('id="2"', 'id="3"').replace('name="Shared removal"', 'name="Remove second"')}</p:spTree>`,
      ),
      slidePart.contentType,
    );
    const second = slide.media.find(({ id }) => id === 3)!;
    const roleCount = slide.relationships.length;

    expect(second.remove()).toBeUndefined();
    expect(slide.media).toEqual([first]);
    expect(slide.relationships).toHaveLength(roleCount);
    expect(pkg.hasPart(mediaUri)).toBe(true);
    expect(pkg.hasPart(posterUri)).toBe(true);
    expect(() => second.name).toThrow(/Shape 3 was not found/);
    expect(() => { second.name = 'Do not recreate'; }).toThrow(/not found/);
    expect(() => second.setTransform({ x: inches(1) })).toThrow(/not found/);
    await expect(second.replacePoster()).rejects.toThrow(/not found/);

    const beforeRemovalRollback = packageSnapshot(pkg);
    expect(() => pkg.transaction(() => {
      first.remove();
      throw new Error('rollback media removal');
    })).toThrow('rollback media removal');
    expect(packageSnapshot(pkg)).toEqual(beforeRemovalRollback);
    expect(slide.media[0]).toBe(first);
    expect(first.name).toBe('Shared removal');

    const hyperlink = pkg.addRelationship(slide.partUri, {
      type: HYPERLINK_RELATIONSHIP,
      target: 'https://example.com/media-details',
      targetMode: 'External',
    });
    const kindRelationship = slide.relationships.find(({ type }) => type.endsWith('/audio'))!;
    pkg.transaction(() => {
      const currentSlide = pkg.requirePart(slide.partUri);
      pkg.setPart(
        slide.partUri,
        new TextDecoder().decode(currentSlide.bytes)
          .replace('r:id=""', `r:id="${hyperlink.id}"`)
          .replace('<a:audioFile ', '<a:videoFile '),
        currentSlide.contentType,
      );
      pkg.updateRelationship(slide.partUri, kindRelationship.id, {
        type: 'http://schemas.microsoft.com/office/2007/relationships/media',
      });
      const mediaPart = pkg.requirePart(mediaUri);
      pkg.setPart(mediaUri, mediaPart.bytes, 'audio/mp3');
    });
    expect(first.mediaPartUri).toBe(mediaUri);
    first.remove();
    expect(slide.media).toHaveLength(0);
    expect(pkg.hasPart(mediaUri)).toBe(false);
    expect(pkg.hasPart(posterUri)).toBe(false);
    expect(slide.relationships.some(({ id }) => id === hyperlink.id)).toBe(false);
    expect(slide.relationships.some(({ type }) =>
      type.endsWith('/audio') || type.endsWith('/media') || type.endsWith('/image'))).toBe(false);

    const external = await slide.addVideo('https://example.com/video.mp4', {
      poster: Uint8Array.of(4),
      posterContentType: 'image/gif',
    });
    const externalPoster = external.posterPartUri!;
    external.remove();
    expect(pkg.hasPart(externalPoster)).toBe(false);
    expect(slide.relationships.some(({ targetMode }) => targetMode === 'External')).toBe(false);
  });

  it('garbage-collects slide media only after the final package reference is deleted', async () => {
    const { pkg, model } = emptyPresentationModel();
    const source = model.addSlide();
    const audio = await source.addAudio(Uint8Array.of(1), {
      contentType: 'audio/mpeg',
      poster: Uint8Array.of(2),
      posterContentType: 'image/png',
    });
    await source.addVideo('https://example.com/video.mp4');
    const image = source.addImage(Uint8Array.of(3), { contentType: 'image/png' });
    const mediaUri = audio.mediaPartUri!;
    const posterUri = audio.posterPartUri!;
    const imageUri = image.sourcePartUri!;
    const duplicate = model.duplicateSlide(0);
    const duplicateMedia = duplicate.media[0]!;
    const keeper = pkg.addRelationship(model.presentationPartUri, {
      type: 'urn:test:retained-media',
      target: relativeRelationshipTarget(model.presentationPartUri, posterUri),
      targetMode: 'Internal',
    });

    model.moveSlide(model.slides.indexOf(duplicate), 0);
    expect(model.slides[0]).toBe(duplicate);
    expect(duplicate.media[0]).toBe(duplicateMedia);
    const beforeRollback = packageSnapshot(pkg);
    expect(() => pkg.transaction(() => {
      model.deleteSlide(model.slides.indexOf(duplicate));
      throw new Error('rollback media slide deletion');
    })).toThrow('rollback media slide deletion');
    expect(packageSnapshot(pkg)).toEqual(beforeRollback);
    expect(model.slides[0]).toBe(duplicate);
    expect(duplicate.media[0]).toBe(duplicateMedia);

    model.deleteSlide(model.slides.indexOf(duplicate));
    expect(pkg.hasPart(mediaUri)).toBe(true);
    expect(pkg.hasPart(posterUri)).toBe(true);
    model.deleteSlide(model.slides.indexOf(source));
    expect(pkg.hasPart(mediaUri)).toBe(false);
    expect(pkg.hasPart(posterUri)).toBe(true);
    expect(pkg.hasPart(imageUri)).toBe(true);
    expect(pkg.relationships(model.presentationPartUri)).toContainEqual(keeper);
  });

  it('creates embedded SVG images atomically before shape-tree extensions', () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const neighbor = slide.addShape('rect', { name: 'Keep neighbor' });
    const slidePart = pkg.requirePart(slide.partUri);
    pkg.setPart(
      slide.partUri,
      new TextDecoder().decode(slidePart.bytes).replace(
        '</p:spTree>',
        '<p:extLst><p:ext uri="urn:keep"><x:keep xmlns:x="urn:keep"/></p:ext>'
          + '</p:extLst></p:spTree>',
      ),
      slidePart.contentType,
    );
    const svgBytes = new Uint8Array([60, 115, 118, 103, 47, 62]);
    const fallbackPngBytes = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
    const image = slide.addSvgImage(svgBytes, fallbackPngBytes, {
      name: 'Vector & <logo>',
      altText: 'SVG & fallback',
      x: inches(1),
      y: inches(2),
      width: inches(3),
      height: inches(2),
      rotation: degrees(45),
      flipHorizontal: true,
      sourceRectangle: { left: 25, top: -10, right: 5, bottom: 0 },
    });

    expect(image).toBeInstanceOf(ImageModel);
    expect(slide.shapes).toEqual([neighbor, image]);
    expect(slide.shapes[1]).toBe(image);
    expect([neighbor.id, image.id]).toEqual([2, 3]);
    expect(image.name).toBe('Vector & <logo>');
    expect(image.transform).toEqual({
      x: inches(1),
      y: inches(2),
      width: inches(3),
      height: inches(2),
      rotation: degrees(45),
      flipHorizontal: true,
      flipVertical: false,
    });
    expect(image.sourceRectangle).toEqual({ left: 25, top: -10, right: 5, bottom: 0 });
    expect(image.sourcePartUri).toBe('/ppt/media/image1.png');

    expect(pkg.parts
      .filter(({ uri }) => uri.startsWith('/ppt/media/'))
      .map(({ uri, contentType, bytes }) => ({ uri, contentType, bytes })))
      .toEqual([
        {
          uri: '/ppt/media/image1.png',
          contentType: 'image/png',
          bytes: new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]),
        },
        {
          uri: '/ppt/media/image1.svg',
          contentType: 'image/svg+xml',
          bytes: new Uint8Array([60, 115, 118, 103, 47, 62]),
        },
      ]);
    const imageRelationships = slide.relationships.filter(
      ({ type }) => type === IMAGE_RELATIONSHIP,
    );
    expect(imageRelationships).toHaveLength(2);
    expect(imageRelationships.map(({ target, targetMode, resolvedTarget }) => ({
      target,
      targetMode,
      resolvedTarget,
    }))).toEqual([
      {
        target: '../media/image1.png',
        targetMode: 'Internal',
        resolvedTarget: '/ppt/media/image1.png',
      },
      {
        target: '../media/image1.svg',
        targetMode: 'Internal',
        resolvedTarget: '/ppt/media/image1.svg',
      },
    ]);

    const source = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(source).toContain(
      '<p:cNvPr id="3" name="Vector &amp; &lt;logo&gt;" descr="SVG &amp; fallback"/>',
    );
    expect(source).toContain(
      `<a:blip r:embed="${imageRelationships[0]!.id}"><a:extLst>`
        + '<a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">'
        + '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" '
        + `r:embed="${imageRelationships[1]!.id}"/></a:ext></a:extLst></a:blip>`,
    );
    expect(source).toContain('<a:srcRect l="25000" t="-10000" r="5000" b="0"/>');
    expect(source.indexOf('name="Keep neighbor"')).toBeLessThan(source.indexOf('name="Vector'));
    expect(source.indexOf('name="Vector')).toBeLessThan(source.indexOf('<p:extLst>'));
    expect(source).toContain('<x:keep xmlns:x="urn:keep"/>');

    svgBytes.fill(0);
    fallbackPngBytes.fill(0);
    expect(pkg.requirePart('/ppt/media/image1.svg').bytes)
      .toEqual(new Uint8Array([60, 115, 118, 103, 47, 62]));
    expect(pkg.requirePart('/ppt/media/image1.png').bytes)
      .toEqual(new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]));
  });

  it('rejects and rolls back embedded SVG image creation without package changes', () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const before = packageSnapshot(pkg);
    const shapes = slide.shapes;
    let reads = 0;
    const accessor = Object.defineProperty({}, 'name', {
      get() {
        reads += 1;
        return 'unsafe';
      },
    });
    const invalidCalls: readonly (() => unknown)[] = [
      () => slide.addSvgImage([] as never, new Uint8Array([2])),
      () => slide.addSvgImage(new Uint8Array(), new Uint8Array([2])),
      () => slide.addSvgImage(new Uint8Array([1]), [] as never),
      () => slide.addSvgImage(new Uint8Array([1]), new Uint8Array()),
      () => slide.addSvgImage(new Uint8Array([1]), new Uint8Array([2]), null as never),
      () => slide.addSvgImage(
        new Uint8Array([1]),
        new Uint8Array([2]),
        Object.create({ name: 'unsafe' }),
      ),
      () => slide.addSvgImage(new Uint8Array([1]), new Uint8Array([2]), accessor),
      () => slide.addSvgImage(new Uint8Array([1]), new Uint8Array([2]), {
        contentType: 'image/svg+xml',
      } as never),
      () => slide.addSvgImage(new Uint8Array([1]), new Uint8Array([2]), {
        width: 0,
      } as never),
      () => slide.addSvgImage(new Uint8Array([1]), new Uint8Array([2]), {
        sourceRectangle: { left: 60, top: 0, right: 40, bottom: 0 },
      }),
    ];
    for (const invoke of invalidCalls) {
      expect(invoke).toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
      expect(slide.shapes).toEqual(shapes);
    }
    expect(reads).toBe(0);

    const malformedSources: readonly ((source: string) => string)[] = [
      (source) => source
        .replace('<p:sld ', '<x:sld xmlns:x="urn:unsafe" ')
        .replace('</p:sld>', '</x:sld>'),
      (source) => source.replace('</p:cSld>', '<p:spTree/></p:cSld>'),
      (source) => source.replace('</p:spTree>', '<p:extLst/><p:extLst/></p:spTree>'),
      (source) => source.replace('id="1"', 'id="not-an-id"'),
    ];
    for (const mutate of malformedSources) {
      const fixture = emptyPresentationModel();
      const malformedSlide = fixture.model.addSlide();
      const part = fixture.pkg.requirePart(malformedSlide.partUri);
      fixture.pkg.setPart(
        malformedSlide.partUri,
        mutate(new TextDecoder().decode(part.bytes)),
        part.contentType,
      );
      const malformedBefore = packageSnapshot(fixture.pkg);
      expect(() => malformedSlide.addSvgImage(
        new Uint8Array([1]),
        new Uint8Array([2]),
      )).toThrow(ModelParseError);
      expect(packageSnapshot(fixture.pkg)).toEqual(malformedBefore);
      expect(malformedSlide.shapes).toEqual([]);
    }

    let rolledBack: ImageModel | undefined;
    expect(() => pkg.transaction(() => {
      rolledBack = slide.addSvgImage(new Uint8Array([1]), new Uint8Array([2]));
      throw new Error('restore embedded SVG image');
    })).toThrow('restore embedded SVG image');
    expect(packageSnapshot(pkg)).toEqual(before);
    expect(slide.shapes).toEqual([]);
    expect(() => rolledBack!.name).toThrow(ModelParseError);

    const created = slide.addSvgImage(new Uint8Array([3]), new Uint8Array([4]));
    expect(created.id).toBe(2);
    expect(created.sourcePartUri).toBe('/ppt/media/image1.png');
    expect(pkg.requirePart('/ppt/media/image1.svg').bytes).toEqual(new Uint8Array([3]));
  });

  it('round-trips embedded SVG image creation in all six presentation formats', async () => {
    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const pkg = await OpcPackage.open(await modelFixture(profile.presentationContentType));
      const model = new PresentationModel(pkg);
      const slide = model.addSlide();
      const image = slide.addSvgImage(
        new Uint8Array([60, 115, 118, 103, 47, 62]),
        new Uint8Array([137, 80, 78, 71]),
        {
          name: `${profile.format} SVG`,
          altText: `${profile.format} vector`,
          x: inches(1),
          width: inches(4),
          height: inches(3),
          sourceRectangle: { left: 12.5, top: 0, right: 12.5, bottom: 0 },
        },
      );
      expect(slide.shapes[0]).toBe(image);
      const fallbackUri = image.sourcePartUri!;
      const svgUri = slide.relationships.find(
        ({ type, resolvedTarget }) =>
          type === IMAGE_RELATIONSHIP && resolvedTarget?.endsWith('.svg'),
      )!.resolvedTarget!;

      const first = new PresentationModel(await OpcPackage.open(await pkg.write()));
      const firstSlide = first.slides.find(({ partUri }) => partUri === slide.partUri)!;
      const firstImage = firstSlide.shapes[0] as ImageModel;
      expect(first.format).toBe(profile.format);
      expect(firstImage).toBeInstanceOf(ImageModel);
      expect(firstImage.name).toBe(`${profile.format} SVG`);
      expect(firstImage.sourcePartUri).toBe(fallbackUri);
      expect(firstImage.sourceRectangle).toEqual({
        left: 12.5,
        top: 0,
        right: 12.5,
        bottom: 0,
      });
      expect(first.opcPackage.requirePart(fallbackUri)).toMatchObject({
        contentType: 'image/png',
        bytes: new Uint8Array([137, 80, 78, 71]),
      });
      expect(first.opcPackage.requirePart(svgUri)).toMatchObject({
        contentType: 'image/svg+xml',
        bytes: new Uint8Array([60, 115, 118, 103, 47, 62]),
      });
      const firstSource = new TextDecoder().decode(
        first.opcPackage.requirePart(firstSlide.partUri).bytes,
      );
      expect(firstSource).toContain(
        '<a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">',
      );
      expect(firstSlide.relationships.filter(({ type }) => type === IMAGE_RELATIONSHIP))
        .toHaveLength(2);

      const second = new PresentationModel(
        await OpcPackage.open(await first.opcPackage.write()),
      );
      const secondSlide = second.slides.find(({ partUri }) => partUri === slide.partUri)!;
      const secondImage = secondSlide.shapes[0] as ImageModel;
      expect(second.format).toBe(profile.format);
      expect(secondImage.name).toBe(`${profile.format} SVG`);
      expect(secondImage.transform).toMatchObject({
        x: inches(1),
        width: inches(4),
        height: inches(3),
      });
      expect(second.opcPackage.requirePart(svgUri).bytes)
        .toEqual(new Uint8Array([60, 115, 118, 103, 47, 62]));
    }
  });

  it('reads and atomically replaces an exclusive SVG pair without changing appearance', async () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const raster = slide.addImage(new Uint8Array([1]), { contentType: 'image/png' });
    const image = slide.addSvgImage(
      new Uint8Array([2, 3]),
      new Uint8Array([4, 5]),
      {
        name: 'Editable vector',
        altText: 'Keep SVG metadata',
        x: inches(1),
        y: inches(2),
        width: inches(3),
        height: inches(4),
        rotation: degrees(15),
        flipVertical: true,
        sourceRectangle: { left: 10, top: 5, right: 20, bottom: -5 },
      },
    );
    const fallbackUri = image.sourcePartUri!;
    const svgUri = slide.relationships.find(
      ({ resolvedTarget }) => resolvedTarget?.endsWith('.svg'),
    )!.resolvedTarget!;
    const part = pkg.requirePart(slide.partUri);
    pkg.setPart(
      slide.partUri,
      new TextDecoder().decode(part.bytes).replace(
        '</a:extLst>',
        '<a:ext uri="urn:keep"><x:keep xmlns:x="urn:keep"/></a:ext></a:extLst>',
      ),
      part.contentType,
    );

    expect(raster.isSvg).toBe(false);
    expect(raster.fallbackPartUri).toBeUndefined();
    expect(raster.svgPartUri).toBeUndefined();
    expect(image.isSvg).toBe(true);
    expect(image.fallbackPartUri).toBe(fallbackUri);
    expect(image.svgPartUri).toBe(svgUri);
    expect(image.sourcePartUri).toBe(fallbackUri);
    const appearance = {
      name: image.name,
      transform: image.transform,
      sourceRectangle: image.sourceRectangle,
    };
    const nextSvg = new Uint8Array([6, 7, 8]);
    const nextFallback = new Uint8Array([9, 10, 11]);

    image.replaceSvgData(nextSvg, nextFallback);

    expect(image.fallbackPartUri).toBe(fallbackUri);
    expect(image.svgPartUri).toBe(svgUri);
    expect(pkg.requirePart(fallbackUri)).toMatchObject({
      contentType: 'image/png',
      bytes: new Uint8Array([9, 10, 11]),
    });
    expect(pkg.requirePart(svgUri)).toMatchObject({
      contentType: 'image/svg+xml',
      bytes: new Uint8Array([6, 7, 8]),
    });
    nextSvg.fill(0);
    nextFallback.fill(0);
    expect(pkg.requirePart(fallbackUri).bytes).toEqual(new Uint8Array([9, 10, 11]));
    expect(pkg.requirePart(svgUri).bytes).toEqual(new Uint8Array([6, 7, 8]));
    expect({
      name: image.name,
      transform: image.transform,
      sourceRectangle: image.sourceRectangle,
    }).toEqual(appearance);
    const source = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(source).toContain('descr="Keep SVG metadata"');
    expect(source).toContain('<x:keep xmlns:x="urn:keep"/>');
    expect(() => image.replaceData(new Uint8Array([12]), 'image/png')).toThrow(
      'Use replaceSvgData() for SVG images',
    );

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedImage = reopened.slides[0]!.shapes[1] as ImageModel;
    expect(reopenedImage.isSvg).toBe(true);
    expect(reopenedImage.sourcePartUri).toBe(fallbackUri);
    expect(reopenedImage.fallbackPartUri).toBe(fallbackUri);
    expect(reopenedImage.svgPartUri).toBe(svgUri);
  });

  it('clones both SVG targets for a duplicate and supports reopen-edit-again', async () => {
    const { pkg, model } = emptyPresentationModel();
    const sourceSlide = model.addSlide();
    const sourceImage = sourceSlide.addSvgImage(
      new Uint8Array([1, 2]),
      new Uint8Array([3, 4]),
      { name: 'Shared vector' },
    );
    const sourceFallbackUri = sourceImage.fallbackPartUri!;
    const sourceSvgUri = sourceImage.svgPartUri!;
    const duplicateSlide = model.duplicateSlide(0);
    const duplicateImage = duplicateSlide.shapes[0] as ImageModel;

    expect(duplicateImage.fallbackPartUri).toBe(sourceFallbackUri);
    expect(duplicateImage.svgPartUri).toBe(sourceSvgUri);
    duplicateImage.replaceSvgData(new Uint8Array([5, 6]), new Uint8Array([7, 8]));
    const duplicateFallbackUri = duplicateImage.fallbackPartUri!;
    const duplicateSvgUri = duplicateImage.svgPartUri!;

    expect(duplicateSlide.shapes[0]).toBe(duplicateImage);
    expect(duplicateFallbackUri).not.toBe(sourceFallbackUri);
    expect(duplicateSvgUri).not.toBe(sourceSvgUri);
    expect(pkg.requirePart(sourceFallbackUri).bytes).toEqual(new Uint8Array([3, 4]));
    expect(pkg.requirePart(sourceSvgUri).bytes).toEqual(new Uint8Array([1, 2]));
    expect(pkg.requirePart(duplicateFallbackUri)).toMatchObject({
      contentType: 'image/png',
      bytes: new Uint8Array([7, 8]),
    });
    expect(pkg.requirePart(duplicateSvgUri)).toMatchObject({
      contentType: 'image/svg+xml',
      bytes: new Uint8Array([5, 6]),
    });

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedDuplicateSlide = reopened.slides.find(
      ({ partUri }) => partUri === duplicateSlide.partUri,
    )!;
    const reopenedDuplicate = reopenedDuplicateSlide.shapes[0] as ImageModel;
    reopenedDuplicate.replaceSvgData(new Uint8Array([9]), new Uint8Array([10]));
    expect(reopenedDuplicate.fallbackPartUri).toBe(duplicateFallbackUri);
    expect(reopenedDuplicate.svgPartUri).toBe(duplicateSvgUri);
    expect(reopened.opcPackage.requirePart(duplicateFallbackUri).bytes)
      .toEqual(new Uint8Array([10]));
    expect(reopened.opcPackage.requirePart(duplicateSvgUri).bytes)
      .toEqual(new Uint8Array([9]));
  });

  it('isolates two SVG pictures that share both relationship ids on one slide', () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const first = slide.addSvgImage(new Uint8Array([1]), new Uint8Array([2]));
    const part = pkg.requirePart(slide.partUri);
    const source = new TextDecoder().decode(part.bytes);
    const picture = source.match(/<p:pic>.*?<\/p:pic>/)?.[0];
    expect(picture).toBeTruthy();
    pkg.setPart(
      slide.partUri,
      source.replace(
        '</p:spTree>',
        `${picture!.replace('id="2"', 'id="3"').replace('name="Image 0"', 'name="Image 1"')}</p:spTree>`,
      ),
      part.contentType,
    );
    const second = slide.shapes.find(({ id }) => id === 3) as ImageModel;
    const sharedFallbackUri = first.fallbackPartUri!;
    const sharedSvgUri = first.svgPartUri!;

    second.replaceSvgData(new Uint8Array([3]), new Uint8Array([4]));

    expect(first.fallbackPartUri).toBe(sharedFallbackUri);
    expect(first.svgPartUri).toBe(sharedSvgUri);
    expect(second.fallbackPartUri).not.toBe(sharedFallbackUri);
    expect(second.svgPartUri).not.toBe(sharedSvgUri);
    expect(pkg.requirePart(sharedFallbackUri).bytes).toEqual(new Uint8Array([2]));
    expect(pkg.requirePart(sharedSvgUri).bytes).toEqual(new Uint8Array([1]));
    expect(slide.relationships.filter(({ type }) => type === IMAGE_RELATIONSHIP))
      .toHaveLength(4);
  });

  it('imports LibreOffice-normalized shared SVG targets and isolates one pair', async () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const created = slide.addSvgImage(new Uint8Array([1]), new Uint8Array([2]));
    const sharedFallbackPartUri = created.fallbackPartUri!;
    const sharedSvgPartUri = created.svgPartUri!;
    pkg.setPart(sharedSvgPartUri, new Uint8Array([1]), 'image/svg');
    const part = pkg.requirePart(slide.partUri);
    const source = new TextDecoder().decode(part.bytes);
    const picture = source.match(/<p:pic>.*?<\/p:pic>/)?.[0];
    expect(picture).toBeTruthy();
    const second = picture!
      .replace('id="2"', 'id="3"')
      .replace('name="Image 0"', 'name="LibreOffice SVG 2"');
    const third = picture!
      .replace('id="2"', 'id="4"')
      .replace('name="Image 0"', 'name="LibreOffice SVG 3"');
    pkg.setPart(
      slide.partUri,
      source.replace('</p:spTree>', `${second}${third}</p:spTree>`),
      part.contentType,
    );
    const [firstImage, editedImage, thirdImage] = slide.shapes as readonly ImageModel[];

    for (const image of [firstImage!, editedImage!, thirdImage!]) {
      expect(image.isSvg).toBe(true);
      expect(image.fallbackPartUri).toBe(sharedFallbackPartUri);
      expect(image.svgPartUri).toBe(sharedSvgPartUri);
    }

    editedImage!.replaceSvgData(new Uint8Array([3]), new Uint8Array([4]));

    expect(firstImage!.fallbackPartUri).toBe(sharedFallbackPartUri);
    expect(firstImage!.svgPartUri).toBe(sharedSvgPartUri);
    expect(thirdImage!.fallbackPartUri).toBe(sharedFallbackPartUri);
    expect(thirdImage!.svgPartUri).toBe(sharedSvgPartUri);
    expect(editedImage!.fallbackPartUri).not.toBe(sharedFallbackPartUri);
    expect(editedImage!.svgPartUri).not.toBe(sharedSvgPartUri);
    expect(pkg.requirePart(sharedFallbackPartUri)).toMatchObject({
      contentType: 'image/png',
      bytes: new Uint8Array([2]),
    });
    expect(pkg.requirePart(sharedSvgPartUri)).toMatchObject({
      contentType: 'image/svg',
      bytes: new Uint8Array([1]),
    });
    expect(pkg.requirePart(editedImage!.fallbackPartUri!)).toMatchObject({
      contentType: 'image/png',
      bytes: new Uint8Array([4]),
    });
    expect(pkg.requirePart(editedImage!.svgPartUri!)).toMatchObject({
      contentType: 'image/svg+xml',
      bytes: new Uint8Array([3]),
    });
    expect(slide.relationships.filter(({ type }) => type === IMAGE_RELATIONSHIP))
      .toHaveLength(4);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedImages = reopened.slides[0]!.shapes as readonly ImageModel[];
    expect(reopenedImages).toHaveLength(3);
    expect(reopenedImages.every(({ isSvg }) => isSvg)).toBe(true);
    expect(reopenedImages[0]!.fallbackPartUri).toBe(sharedFallbackPartUri);
    expect(reopenedImages[2]!.svgPartUri).toBe(sharedSvgPartUri);
    expect(reopenedImages[1]!.fallbackPartUri).toBe(editedImage!.fallbackPartUri);
    expect(reopenedImages[1]!.svgPartUri).toBe(editedImage!.svgPartUri);
  });

  it('handles independently shared and noncanonical SVG targets', () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const image = slide.addSvgImage(new Uint8Array([1]), new Uint8Array([2]));
    const neighbor = model.addSlide();
    const originalFallbackUri = image.fallbackPartUri!;
    const originalSvgUri = image.svgPartUri!;
    pkg.addRelationship(neighbor.partUri, {
      type: IMAGE_RELATIONSHIP,
      target: '../media/image1.png',
      targetMode: 'Internal',
    });

    image.replaceSvgData(new Uint8Array([3]), new Uint8Array([4]));

    expect(image.fallbackPartUri).not.toBe(originalFallbackUri);
    expect(image.svgPartUri).toBe(originalSvgUri);
    expect(pkg.requirePart(originalFallbackUri).bytes).toEqual(new Uint8Array([2]));
    expect(pkg.requirePart(originalSvgUri).bytes).toEqual(new Uint8Array([3]));

    const fallbackRelationship = slide.relationships.find(
      ({ resolvedTarget }) => resolvedTarget === image.fallbackPartUri,
    )!;
    pkg.setPart('/ppt/media/noncanonical.bin', new Uint8Array([5]), 'application/octet-stream');
    pkg.updateRelationship(slide.partUri, fallbackRelationship.id, {
      target: '../media/noncanonical.bin',
    });
    pkg.setPart(originalSvgUri, new Uint8Array([6]), 'image/svg');
    expect(image.fallbackPartUri).toBe('/ppt/media/noncanonical.bin');
    expect(image.svgPartUri).toBe(originalSvgUri);

    image.replaceSvgData(new Uint8Array([7]), new Uint8Array([8]));

    expect(image.fallbackPartUri).toMatch(/\.png$/);
    expect(image.fallbackPartUri).not.toBe('/ppt/media/noncanonical.bin');
    expect(image.svgPartUri).toBe(originalSvgUri);
    expect(pkg.requirePart(image.fallbackPartUri!)).toMatchObject({
      contentType: 'image/png',
      bytes: new Uint8Array([8]),
    });
    expect(pkg.requirePart(originalSvgUri)).toMatchObject({
      contentType: 'image/svg+xml',
      bytes: new Uint8Array([7]),
    });
    expect(pkg.requirePart('/ppt/media/noncanonical.bin').bytes)
      .toEqual(new Uint8Array([5]));
  });

  it('rejects malformed SVG state and rolls paired replacements back exactly', () => {
    const malformedFixture = emptyPresentationModel();
    const malformedSlide = malformedFixture.model.addSlide();
    const malformedImage = malformedSlide.addSvgImage(
      new Uint8Array([1]),
      new Uint8Array([2]),
    );
    const svgRelationship = malformedSlide.relationships.find(
      ({ resolvedTarget }) => resolvedTarget?.endsWith('.svg'),
    )!;
    malformedFixture.pkg.updateRelationship(malformedSlide.partUri, svgRelationship.id, {
      type: 'urn:wrong',
    });
    const malformedBefore = packageSnapshot(malformedFixture.pkg);

    expect(malformedImage.isSvg).toBe(false);
    expect(malformedImage.fallbackPartUri).toBeUndefined();
    expect(malformedImage.svgPartUri).toBeUndefined();
    expect(() => malformedImage.replaceSvgData(
      new Uint8Array([3]),
      new Uint8Array([4]),
    )).toThrow('Image 2 is not a safely editable SVG image');
    expect(() => malformedImage.replaceData(new Uint8Array([5]), 'image/png'))
      .toThrow('Use replaceSvgData() for SVG images');
    expect(packageSnapshot(malformedFixture.pkg)).toEqual(malformedBefore);

    const rollbackFixture = emptyPresentationModel();
    const rollbackSlide = rollbackFixture.model.addSlide();
    const rollbackImage = rollbackSlide.addSvgImage(
      new Uint8Array([6]),
      new Uint8Array([7]),
    );
    const rollbackBefore = packageSnapshot(rollbackFixture.pkg);
    expect(() => rollbackFixture.pkg.transaction(() => {
      rollbackImage.replaceSvgData(new Uint8Array([8]), new Uint8Array([9]));
      throw new Error('restore SVG pair');
    })).toThrow('restore SVG pair');
    expect(packageSnapshot(rollbackFixture.pkg)).toEqual(rollbackBefore);
    expect(rollbackImage.isSvg).toBe(true);

    const failureFixture = emptyPresentationModel();
    const failureSlide = failureFixture.model.addSlide();
    const failureImage = failureSlide.addSvgImage(
      new Uint8Array([10]),
      new Uint8Array([11]),
    );
    const failureSvgUri = failureImage.svgPartUri!;
    failureFixture.pkg.setPart(failureSvgUri, new Uint8Array([10]), 'image/svg');
    failureFixture.pkg.setPart('/[Content_Types].xml', '<invalid/>');
    const failureBefore = packageSnapshot(failureFixture.pkg);
    for (const invoke of [
      () => failureImage.replaceSvgData([] as never, new Uint8Array([12])),
      () => failureImage.replaceSvgData(new Uint8Array([12]), new Uint8Array()),
    ]) {
      expect(invoke).toThrow();
      expect(packageSnapshot(failureFixture.pkg)).toEqual(failureBefore);
    }
    expect(() => failureImage.replaceSvgData(
      new Uint8Array([12]),
      new Uint8Array([13]),
    )).toThrow(/Invalid \[Content_Types\]\.xml/);
    expect(packageSnapshot(failureFixture.pkg)).toEqual(failureBefore);
  });

  it('creates multiple embedded raster images before extensions without disturbing neighbors', () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const preset = slide.addShape('rect', { name: 'Keep shape' });
    const slidePart = pkg.requirePart(slide.partUri);
    pkg.setPart(
      slide.partUri,
      new TextDecoder().decode(slidePart.bytes).replace(
        '</p:spTree>',
        '<p:extLst><p:ext uri="urn:keep"><x:keep xmlns:x="urn:keep"/></p:ext>' +
        '</p:extLst></p:spTree>',
      ),
      slidePart.contentType,
    );
    pkg.setPart('/ppt/custom/keep.bin', new Uint8Array([9, 8, 7]), 'application/octet-stream');
    const keepRelationship = pkg.addRelationship(slide.partUri, {
      type: 'urn:keep:relationship',
      target: '../custom/keep.bin',
      targetMode: 'Internal',
    });

    const png = slide.addImage(new Uint8Array([1]), { contentType: 'image/png' });
    const jpeg = slide.addImage(new Uint8Array([2]), { contentType: 'image/jpeg' });
    const gif = slide.addImage(new Uint8Array([3]), { contentType: 'image/gif' });

    expect(slide.shapes).toEqual([preset, png, jpeg, gif]);
    expect(slide.shapes.map(({ id }) => id)).toEqual([2, 3, 4, 5]);
    expect([png.name, jpeg.name, gif.name]).toEqual(['Image 0', 'Image 1', 'Image 2']);
    for (const image of [png, jpeg, gif]) {
      expect(image.transform).toEqual({
        x: 0,
        y: 0,
        width: 914_400,
        height: 914_400,
        rotation: 0,
        flipHorizontal: false,
        flipVertical: false,
      });
    }
    expect([png.sourcePartUri, jpeg.sourcePartUri, gif.sourcePartUri]).toEqual([
      '/ppt/media/image1.png',
      '/ppt/media/image1.jpeg',
      '/ppt/media/image1.gif',
    ]);
    expect(pkg.parts
      .filter(({ uri }) => uri.startsWith('/ppt/media/'))
      .map(({ uri, contentType, bytes }) => ({ uri, contentType, bytes })))
      .toEqual([
        { uri: '/ppt/media/image1.png', contentType: 'image/png', bytes: new Uint8Array([1]) },
        { uri: '/ppt/media/image1.jpeg', contentType: 'image/jpeg', bytes: new Uint8Array([2]) },
        { uri: '/ppt/media/image1.gif', contentType: 'image/gif', bytes: new Uint8Array([3]) },
      ]);
    expect(slide.relationships.find(({ id }) => id === keepRelationship.id))
      .toEqual(keepRelationship);
    expect(pkg.requirePart('/ppt/custom/keep.bin').bytes).toEqual(new Uint8Array([9, 8, 7]));
    expect(slide.relationships
      .filter(({ type }) => type === IMAGE_RELATIONSHIP)
      .map(({ target }) => target))
      .toEqual(['../media/image1.png', '../media/image1.jpeg', '../media/image1.gif']);

    const xml = LosslessXmlDocument.parse(pkg.requirePart(slide.partUri).bytes);
    const shapeTree = xml.elements('spTree')[0]!;
    expect(shapeTree.children
      .filter((child) => child.type === 'element')
      .map(({ localName }) => localName)
      .filter((name) => name === 'sp' || name === 'pic' || name === 'extLst'))
      .toEqual(['sp', 'pic', 'pic', 'pic', 'extLst']);
    expect(new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes))
      .toContain('<x:keep xmlns:x="urn:keep"/>');
  });

  it('rejects invalid embedded raster image inputs without changing package or shape state', () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const before = packageSnapshot(pkg);
    const shapes = slide.shapes;
    let reads = 0;
    const accessor = Object.defineProperty({ contentType: 'image/png' }, 'name', {
      get() {
        reads += 1;
        return 'unsafe';
      },
    });
    const invalidCalls: readonly (() => unknown)[] = [
      () => slide.addImage([] as never, { contentType: 'image/png' }),
      () => slide.addImage(new Uint8Array(), { contentType: 'image/png' }),
      () => slide.addImage(new Uint8Array([1]), null as never),
      () => slide.addImage(new Uint8Array([1]), Object.create({ contentType: 'image/png' })),
      () => slide.addImage(new Uint8Array([1]), accessor as never),
      () => slide.addImage(new Uint8Array([1]), {
        contentType: 'image/png',
        [Symbol('unsafe')]: true,
      } as never),
      () => slide.addImage(new Uint8Array([1]), {
        contentType: 'image/png',
        unknown: true,
      } as never),
      () => slide.addImage(new Uint8Array([1]), {} as never),
      () => slide.addImage(new Uint8Array([1]), { contentType: undefined } as never),
      () => slide.addImage(new Uint8Array([1]), { contentType: 'image/svg+xml' } as never),
      () => slide.addImage(new Uint8Array([1]), {
        contentType: 'image/png',
        name: 'bad\u0000name',
      }),
      () => slide.addImage(new Uint8Array([1]), {
        contentType: 'image/png',
        altText: 1,
      } as never),
      () => slide.addImage(
        new Uint8Array([1]),
        { contentType: 'image/png', x: Number.NaN } as never,
      ),
      () => slide.addImage(
        new Uint8Array([1]),
        { contentType: 'image/png', width: 0 } as never,
      ),
      () => slide.addImage(new Uint8Array([1]), {
        contentType: 'image/png',
        rotation: 21_600_001,
      } as never),
      () => slide.addImage(new Uint8Array([1]), {
        contentType: 'image/png',
        flipVertical: 1,
      } as never),
      () => slide.addImage(new Uint8Array([1]), {
        contentType: 'image/png',
        sourceRectangle: { left: 60, top: 0, right: 40, bottom: 0 },
      }),
    ];

    for (const invoke of invalidCalls) {
      expect(invoke).toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
      expect(slide.shapes).toEqual(shapes);
    }
    expect(reads).toBe(0);
  });

  it('rolls back embedded raster image resources for malformed slides and outer transactions', () => {
    const malformedSources: readonly ((source: string) => string)[] = [
      (source) => source
        .replace('<p:sld ', '<x:sld xmlns:x="urn:unsafe" ')
        .replace('</p:sld>', '</x:sld>'),
      (source) => source.replace('</p:cSld>', '<p:spTree/></p:cSld>'),
      (source) => source.replace(
        '</p:spTree>',
        '<p:extLst/><p:extLst/></p:spTree>',
      ),
      (source) => source.replace('</p:spTree>', '<p:cNvPr id="1"/></p:spTree>'),
      (source) => source.replace('id="1"', 'id="not-an-id"'),
      (source) => source.replace('id="1"', 'id="4294967296"'),
      (source) => source.replace('id="1"', 'id="4294967295"'),
    ];

    for (const mutate of malformedSources) {
      const { pkg, model } = emptyPresentationModel();
      const slide = model.addSlide();
      const part = pkg.requirePart(slide.partUri);
      pkg.setPart(
        slide.partUri,
        mutate(new TextDecoder().decode(part.bytes)),
        part.contentType,
      );
      const before = packageSnapshot(pkg);
      const shapes = slide.shapes;
      expect(() => slide.addImage(
        new Uint8Array([1]),
        { contentType: 'image/png' },
      )).toThrow(ModelParseError);
      expect(packageSnapshot(pkg)).toEqual(before);
      expect(slide.shapes).toEqual(shapes);
    }

    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const before = packageSnapshot(pkg);
    let rolledBack: ImageModel | undefined;
    expect(() => pkg.transaction(() => {
      rolledBack = slide.addImage(new Uint8Array([1]), { contentType: 'image/png' });
      throw new Error('restore embedded raster image');
    })).toThrow('restore embedded raster image');
    expect(packageSnapshot(pkg)).toEqual(before);
    expect(slide.shapes).toEqual([]);
    expect(() => rolledBack!.name).toThrow(ModelParseError);
    const created = slide.addImage(new Uint8Array([2]), { contentType: 'image/png' });
    expect(created.id).toBe(2);
    expect(created.sourcePartUri).toBe('/ppt/media/image1.png');
  });

  it('reads, edits, clears, rolls back, and duplicates image source rectangles', async () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const image = slide.addImage(new Uint8Array([1, 2, 3]), {
      contentType: 'image/png',
      sourceRectangle: { left: 25, top: -10, right: 5, bottom: 0 },
    });
    const first = image.sourceRectangle;
    const second = image.sourceRectangle;
    expect(first).toEqual({ left: 25, top: -10, right: 5, bottom: 0 });
    expect(second).toEqual(first);
    expect(second).not.toBe(first);
    expect(Object.isFrozen(first)).toBe(true);

    const beforeNoOp = pkg.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...pkg.mutations];
    image.sourceRectangle = { left: 25, top: -10, right: 5, bottom: 0 };
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeNoOp);
    expect(pkg.mutations).toEqual(noOpJournal);

    image.sourceRectangle = { left: 10, top: 20, right: 30, bottom: 0 };
    expect(image.sourceRectangle).toEqual({ left: 10, top: 20, right: 30, bottom: 0 });
    expect(slide.shapes[0]).toBe(image);
    expect(pkg.requirePart(image.sourcePartUri!).bytes).toEqual(new Uint8Array([1, 2, 3]));

    const rollbackBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const rollbackJournal = [...pkg.mutations];
    expect(() => pkg.transaction(() => {
      image.sourceRectangle = { left: -25, top: 0, right: -25, bottom: 0 };
      throw new Error('restore image source rectangle');
    })).toThrow('restore image source rectangle');
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(rollbackBytes);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(image.sourceRectangle).toEqual({ left: 10, top: 20, right: 30, bottom: 0 });

    const duplicate = model.duplicateSlide(0);
    const duplicateImage = duplicate.shapes[0] as ImageModel;
    expect(duplicateImage).toBeInstanceOf(ImageModel);
    expect(duplicateImage.sourceRectangle).toEqual(image.sourceRectangle);
    duplicateImage.sourceRectangle = { left: -25, top: 0, right: -25, bottom: 0 };
    expect(duplicateImage.sourceRectangle).toEqual({
      left: -25,
      top: 0,
      right: -25,
      bottom: 0,
    });
    expect(image.sourceRectangle).toEqual({ left: 10, top: 20, right: 30, bottom: 0 });
    expect(duplicateImage.sourcePartUri).toBe(image.sourcePartUri);

    image.sourceRectangle = undefined;
    expect(image.sourceRectangle).toBeUndefined();
    const clearBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const clearJournal = [...pkg.mutations];
    image.sourceRectangle = undefined;
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(clearBytes);
    expect(pkg.mutations).toEqual(clearJournal);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSource = reopened.slides[0]!.shapes[0] as ImageModel;
    const reopenedDuplicate = reopened.slides[1]!.shapes[0] as ImageModel;
    expect(reopenedSource.sourceRectangle).toBeUndefined();
    expect(reopenedDuplicate.sourceRectangle).toEqual({
      left: -25,
      top: 0,
      right: -25,
      bottom: 0,
    });
  });

  it('repairs one malformed image source rectangle and rejects ambiguous ownership', () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const image = slide.addImage(new Uint8Array([1, 2, 3]), {
      contentType: 'image/png',
    });
    const mediaPartUri = image.sourcePartUri!;
    const relationships = slide.relationships;
    const part = pkg.requirePart(slide.partUri);
    pkg.setPart(
      slide.partUri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:stretch>',
        '<a:srcRect l="bad" custom="remove"/><x:keep xmlns:x="urn:keep"/>' +
        '<a:stretch>',
      ),
      part.contentType,
    );
    expect(image.sourceRectangle).toBeUndefined();

    image.sourceRectangle = { left: 25, top: -10, right: 5, bottom: 0 };
    expect(image.sourceRectangle).toEqual({ left: 25, top: -10, right: 5, bottom: 0 });
    let source = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(source).toContain(
      '<a:srcRect l="25000" t="-10000" r="5000" b="0"/>',
    );
    expect(source).not.toContain('custom="remove"');
    expect(source).toContain('<x:keep xmlns:x="urn:keep"/>');
    expect(image.sourcePartUri).toBe(mediaPartUri);
    expect(slide.relationships).toEqual(relationships);
    expect(pkg.requirePart(mediaPartUri).bytes).toEqual(new Uint8Array([1, 2, 3]));

    const current = '<a:srcRect l="25000" t="-10000" r="5000" b="0"/>';
    const duplicate = `${current}<a:srcRect l="0" t="0" r="0" b="0"/>`;
    const updated = pkg.requirePart(slide.partUri);
    source = new TextDecoder().decode(updated.bytes).replace(current, duplicate);
    pkg.setPart(slide.partUri, source, updated.contentType);
    expect(image.sourceRectangle).toBeUndefined();
    const before = packageSnapshot(pkg);
    expect(() => {
      image.sourceRectangle = { left: 0, top: 0, right: 0, bottom: 0 };
    }).toThrow(ModelParseError);
    expect(() => {
      image.sourceRectangle = undefined;
    }).toThrow(ModelParseError);
    expect(packageSnapshot(pkg)).toEqual(before);
    expect(slide.shapes[0]).toBe(image);
  });

  it('preserves embedded raster image edits and clone-on-write state through reopen', async () => {
    const { pkg, model } = emptyPresentationModel();
    const sourceSlide = model.addSlide();
    const sourceImage = sourceSlide.addImage(new Uint8Array([1, 2, 3]), {
      contentType: 'image/png',
      name: 'Lifecycle image',
      altText: 'Lifecycle alt text',
    });
    const originalUri = sourceImage.sourcePartUri!;

    sourceImage.setTransform({
      x: inches(2),
      y: inches(1),
      width: inches(4),
      height: inches(3),
      rotation: degrees(30),
      flipVertical: true,
    });
    sourceImage.replaceData(new Uint8Array([4, 5, 6]), 'image/png');
    expect(sourceImage.sourcePartUri).toBe(originalUri);
    expect(sourceSlide.shapes[0]).toBe(sourceImage);
    expect(pkg.requirePart(originalUri)).toMatchObject({
      contentType: 'image/png',
      bytes: new Uint8Array([4, 5, 6]),
    });

    const duplicateSlide = model.duplicateSlide(model.slides.indexOf(sourceSlide));
    const duplicateImage = duplicateSlide.shapes[0] as ImageModel;
    expect(duplicateImage).toBeInstanceOf(ImageModel);
    expect(duplicateImage.sourcePartUri).toBe(originalUri);
    expect(duplicateImage.name).toBe('Lifecycle image');
    expect(duplicateImage.transform).toEqual(sourceImage.transform);

    duplicateImage.replaceData(new Uint8Array([7, 8, 9]), 'image/png');
    const duplicateUri = duplicateImage.sourcePartUri!;
    expect(duplicateUri).not.toBe(originalUri);
    expect(duplicateSlide.shapes[0]).toBe(duplicateImage);
    expect(pkg.requirePart(originalUri).bytes).toEqual(new Uint8Array([4, 5, 6]));
    expect(pkg.requirePart(duplicateUri)).toMatchObject({
      contentType: 'image/png',
      bytes: new Uint8Array([7, 8, 9]),
    });

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSourceSlide = reopened.slides.find(
      ({ partUri }) => partUri === sourceSlide.partUri,
    )!;
    const reopenedDuplicateSlide = reopened.slides.find(
      ({ partUri }) => partUri === duplicateSlide.partUri,
    )!;
    const reopenedSource = reopenedSourceSlide.shapes[0] as ImageModel;
    const reopenedDuplicate = reopenedDuplicateSlide.shapes[0] as ImageModel;
    expect(reopenedSource).toBeInstanceOf(ImageModel);
    expect(reopenedDuplicate).toBeInstanceOf(ImageModel);
    expect(reopenedSourceSlide.shapes[0]).toBe(reopenedSource);
    expect(reopenedDuplicateSlide.shapes[0]).toBe(reopenedDuplicate);
    expect([reopenedSource.name, reopenedDuplicate.name]).toEqual([
      'Lifecycle image',
      'Lifecycle image',
    ]);
    expect(reopenedSource.transform).toEqual({
      x: inches(2),
      y: inches(1),
      width: inches(4),
      height: inches(3),
      rotation: degrees(30),
      flipHorizontal: false,
      flipVertical: true,
    });
    expect(reopenedDuplicate.transform).toEqual(reopenedSource.transform);
    expect(reopenedSource.sourcePartUri).toBe(originalUri);
    expect(reopenedDuplicate.sourcePartUri).toBe(duplicateUri);
    expect(reopened.opcPackage.requirePart(originalUri)).toMatchObject({
      contentType: 'image/png',
      bytes: new Uint8Array([4, 5, 6]),
    });
    expect(reopened.opcPackage.requirePart(duplicateUri)).toMatchObject({
      contentType: 'image/png',
      bytes: new Uint8Array([7, 8, 9]),
    });
    for (const [slide, uri] of [
      [reopenedSourceSlide, originalUri],
      [reopenedDuplicateSlide, duplicateUri],
    ] as const) {
      expect(slide.relationships.some(({ type, targetMode, resolvedTarget }) =>
        type === IMAGE_RELATIONSHIP
        && targetMode === 'Internal'
        && resolvedTarget === uri)).toBe(true);
    }
  });

  it('rolls back, moves, and deletes embedded raster image slides by shared policy', () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const before = packageSnapshot(pkg);
    let rolledBack: ImageModel | undefined;
    expect(() => pkg.transaction(() => {
      rolledBack = slide.addImage(new Uint8Array([1]), {
        contentType: 'image/png',
        name: 'Temporary image',
      });
      rolledBack.setTransform({ x: inches(3), rotation: degrees(15) });
      rolledBack.replaceData(new Uint8Array([2, 3]), 'image/png');
      throw new Error('restore complete image lifecycle');
    })).toThrow('restore complete image lifecycle');
    expect(packageSnapshot(pkg)).toEqual(before);
    expect(slide.shapes).toEqual([]);
    expect(() => rolledBack!.name).toThrow(ModelParseError);

    const image = slide.addImage(new Uint8Array([4, 5]), {
      contentType: 'image/png',
      name: 'Persistent image',
    });
    const mediaPartUri = image.sourcePartUri!;
    const slidePartUri = slide.partUri;
    const slideRelationshipId = slide.relationshipId;
    const neighbor = model.addSlide();
    model.moveSlide(model.slides.indexOf(slide), model.slides.length - 1);
    expect(model.slides).toEqual([neighbor, slide]);
    expect(model.slides[1]).toBe(slide);
    expect(slide.shapes[0]).toBe(image);
    expect(image.sourcePartUri).toBe(mediaPartUri);

    model.deleteSlide(model.slides.indexOf(slide));
    expect(model.slides).toEqual([neighbor]);
    expect(pkg.hasPart(slidePartUri)).toBe(false);
    expect(pkg.relationships(model.presentationPartUri)
      .some(({ id }) => id === slideRelationshipId)).toBe(false);
    expect(pkg.hasPart(mediaPartUri)).toBe(true);
    expect(pkg.requirePart(mediaPartUri)).toMatchObject({
      contentType: 'image/png',
      bytes: new Uint8Array([4, 5]),
    });
    expect(pkg.graph.find(({ uri }) => uri === mediaPartUri)?.incoming).toEqual([]);
    expect(() => image.name).toThrow(/Missing package part/);
  });

  it('round-trips embedded raster image lifecycle in all six presentation formats', async () => {
    const definitions = [
      { contentType: 'image/png' as const, bytes: new Uint8Array([1]), name: 'PNG image' },
      { contentType: 'image/jpeg' as const, bytes: new Uint8Array([2]), name: 'JPEG image' },
      { contentType: 'image/gif' as const, bytes: new Uint8Array([3]), name: 'GIF image' },
    ];

    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const pkg = await OpcPackage.open(await modelFixture(profile.presentationContentType));
      const model = new PresentationModel(pkg);
      const slide = model.addSlide();
      for (const definition of definitions) {
        slide.addImage(definition.bytes, {
          contentType: definition.contentType,
          name: definition.name,
          altText: `${definition.name} description`,
          ...(definition.contentType === 'image/png'
            ? { sourceRectangle: { left: 25, top: -10, right: 5, bottom: 0 } }
            : {}),
        });
      }

      const first = new PresentationModel(await OpcPackage.open(await pkg.write()));
      const firstSlide = first.slides.find(({ partUri }) => partUri === slide.partUri)!;
      const firstImages = firstSlide.shapes.filter(
        (shape): shape is ImageModel => shape instanceof ImageModel,
      );
      expect(first.format).toBe(profile.format);
      expect(firstImages.map(({ name }) => name)).toEqual(definitions.map(({ name }) => name));
      expect(firstSlide.shapes.filter(
        (shape): shape is ImageModel => shape instanceof ImageModel,
      )).toEqual(firstImages);
      const firstSlideXml = new TextDecoder().decode(
        first.opcPackage.requirePart(firstSlide.partUri).bytes,
      );
      expect(firstSlideXml).toContain(
        '<a:srcRect l="25000" t="-10000" r="5000" b="0"/>',
      );
      for (const [index, image] of firstImages.entries()) {
        const definition = definitions[index]!;
        const uri = image.sourcePartUri!;
        expect(first.opcPackage.requirePart(uri)).toMatchObject({
          contentType: definition.contentType,
          bytes: definition.bytes,
        });
        expect(firstSlide.relationships.some(({ type, targetMode, resolvedTarget }) =>
          type === IMAGE_RELATIONSHIP
          && targetMode === 'Internal'
          && resolvedTarget === uri)).toBe(true);
        image.setTransform({
          x: inches(index + 1),
          y: inches(index + 2),
          rotation: degrees((index + 1) * 15),
          flipHorizontal: index === 1,
        });
        image.replaceData(
          new Uint8Array([10 + index, 20 + index]),
          definition.contentType,
        );
      }

      const second = new PresentationModel(
        await OpcPackage.open(await first.opcPackage.write()),
      );
      const secondSlide = second.slides.find(({ partUri }) => partUri === slide.partUri)!;
      const secondImages = secondSlide.shapes.filter(
        (shape): shape is ImageModel => shape instanceof ImageModel,
      );
      expect(second.format).toBe(profile.format);
      expect(secondImages.map(({ name }) => name)).toEqual(definitions.map(({ name }) => name));
      expect(secondSlide.shapes.filter(
        (shape): shape is ImageModel => shape instanceof ImageModel,
      )).toEqual(secondImages);
      for (const [index, image] of secondImages.entries()) {
        const uri = image.sourcePartUri!;
        expect(image.transform).toEqual({
          x: inches(index + 1),
          y: inches(index + 2),
          width: inches(1),
          height: inches(1),
          rotation: degrees((index + 1) * 15),
          flipHorizontal: index === 1,
          flipVertical: false,
        });
        expect(second.opcPackage.requirePart(uri)).toMatchObject({
          contentType: definitions[index]!.contentType,
          bytes: new Uint8Array([10 + index, 20 + index]),
        });
        expect(secondSlide.relationships.some(({ type, targetMode, resolvedTarget }) =>
          type === IMAGE_RELATIONSHIP
          && targetMode === 'Internal'
          && resolvedTarget === uri)).toBe(true);
      }
    }
  });

  it('creates detached styled, multi-path, and empty custom shapes before extensions', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const part = pkg.requirePart(slide.partUri);
    pkg.setPart(
      slide.partUri,
      new TextDecoder().decode(part.bytes).replace(
        '</p:spTree>',
        '<p:extLst><p:ext uri="urn:test"><x:keep xmlns:x="urn:test"/>' +
        '</p:ext></p:extLst></p:spTree>',
      ),
      part.contentType,
    );
    const mutableGeometry = structuredClone(customTriangleGeometry) as {
      paths: Array<{
        width: number;
        height: number;
        commands: Array<CustomGeometry['paths'][number]['commands'][number]>;
      }>;
    };
    const options = {
      name: 'Custom triangle',
      x: inches(1),
      y: inches(1),
      width: inches(4),
      height: inches(3),
      fill: { kind: 'solid' as const, color: { kind: 'scheme' as const, value: 'accent1' as const } },
      line: {
        kind: 'line' as const,
        color: { kind: 'srgb' as const, value: '112233' },
        width: 2,
      },
      arrows: { end: 'triangle' as const },
      shadow: { kind: 'outer' as const },
      hyperlink: { url: 'https://example.com/custom' },
    };
    const custom = slide.addCustomShape(mutableGeometry, options);
    mutableGeometry.paths[0]!.width = 1;
    mutableGeometry.paths[0]!.commands.splice(0);
    options.name = 'Changed';
    options.line.color.value = 'FFFFFF';

    const multiPath: CustomGeometry = {
      paths: [
        {
          width: 100,
          height: 100,
          fill: 'none',
          stroke: false,
          commands: [{ kind: 'moveTo', point: { x: 0, y: 0 } }],
        },
        { width: 200, height: 300, extrusionOk: false, commands: [] },
      ],
    };
    const multi = slide.addCustomShape(multiPath, { name: 'Multiple paths' });
    const emptyGeometry: CustomGeometry = {
      paths: [{ width: 1, height: 1, commands: [] }],
    };
    const empty = slide.addCustomShape(emptyGeometry, { name: 'Empty path' });

    expect([custom.id, multi.id, empty.id]).toEqual([2, 3, 4]);
    expect(slide.shapes).toEqual([custom, multi, empty]);
    expect(slide.shapes[0]).toBe(custom);
    expect(custom.kind).toBe('shape');
    expect(custom.presetType).toBeUndefined();
    expect(custom.name).toBe('Custom triangle');
    expect(custom.transform).toEqual({
      x: inches(1),
      y: inches(1),
      width: inches(4),
      height: inches(3),
      rotation: 0,
      flipHorizontal: false,
      flipVertical: false,
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
    expect(custom.hyperlink).toEqual({ url: 'https://example.com/custom' });
    expect(readCreatedCustomGeometry(model, slide, custom.id)).toEqual(customTriangleGeometry);
    expect(readCreatedCustomGeometry(model, slide, multi.id)).toEqual(multiPath);
    expect(readCreatedCustomGeometry(model, slide, empty.id)).toEqual(emptyGeometry);

    const xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml.indexOf('<a:custGeom>')).toBeLessThan(xml.indexOf('<p:extLst>'));
    expect(xml).toContain('<x:keep xmlns:x="urn:test"/>');
  });

  it('rejects invalid custom shape inputs and targets without package or ID mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const before = packageSnapshot(pkg);
    const shapes = slide.shapes;
    let calls = 0;
    const optionsAccessor = Object.defineProperty({}, 'name', {
      enumerable: true,
      get() {
        calls += 1;
        return 'Unsafe';
      },
    });
    const pointAccessor = Object.defineProperty({ y: 0 }, 'x', {
      enumerable: true,
      get() {
        calls += 1;
        return 0;
      },
    });

    for (const operation of [
      () => slide.addCustomShape({ paths: [] }),
      () => slide.addCustomShape({ paths: [{ width: 0, height: 1, commands: [] }] }),
      () => slide.addCustomShape({
        paths: [{
          width: 1,
          height: 1,
          commands: [{ kind: 'lineTo', point: { x: 0, y: 0 } }],
        }],
      }),
      () => slide.addCustomShape({
        paths: [{
          width: 1,
          height: 1,
          commands: [{ kind: 'moveTo', point: pointAccessor as never }],
        }],
      }),
      () => slide.addCustomShape(customTriangleGeometry, optionsAccessor),
      () => slide.addCustomShape(customTriangleGeometry, { adjustments: undefined } as never),
      () => slide.addCustomShape(customTriangleGeometry, { hyperlink: { slide: 99 } }),
    ]) {
      expect(operation).toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
      expect(slide.shapes).toEqual(shapes);
    }
    expect(calls).toBe(0);
    expect(slide.addCustomShape(customTriangleGeometry).id).toBe(2);
  });

  it('rolls back, duplicates, moves, deletes, and reopens custom shapes in all six formats', async () => {
    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const pkg = await OpcPackage.open(await modelFixture(profile.presentationContentType));
      const model = new PresentationModel(pkg);
      const slide = model.addSlide();
      const before = packageSnapshot(pkg);
      let rolledBack: ShapeModel | undefined;
      expect(() => pkg.transaction(() => {
        rolledBack = slide.addCustomShape(customTriangleGeometry, {
          hyperlink: { url: `https://example.com/rollback/${profile.format}` },
        });
        throw new Error('restore custom shape creation');
      })).toThrow('restore custom shape creation');
      expect(packageSnapshot(pkg)).toEqual(before);
      expect(() => rolledBack!.name).toThrow(ModelParseError);

      const created = slide.addCustomShape(customTriangleGeometry, {
        name: `Custom ${profile.format}`,
        fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
        hyperlink: { url: `https://example.com/${profile.format}` },
      });
      const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
      expect(readCreatedCustomGeometry(model, duplicate, created.id)).toEqual(
        customTriangleGeometry,
      );
      created.setTransform({ x: inches(2) });
      created.fill = { kind: 'solid', color: { kind: 'srgb', value: '445566' } };
      expect(readCreatedCustomGeometry(model, slide, created.id)).toEqual(customTriangleGeometry);
      expect(readCreatedCustomGeometry(model, duplicate, created.id)).toEqual(
        customTriangleGeometry,
      );
      model.moveSlide(model.slides.indexOf(duplicate), 0);
      model.deleteSlide(model.slides.indexOf(duplicate));

      const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
      const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!;
      const reopenedShape = reopenedSlide.shapes.find(({ id }) => id === created.id) as ShapeModel;
      expect(reopened.format).toBe(profile.format);
      expect(reopenedShape.name).toBe(`Custom ${profile.format}`);
      expect(reopenedShape.transform.x).toBe(inches(2));
      expect(reopenedShape.fill).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: '445566' },
      });
      expect(readCreatedCustomGeometry(reopened, reopenedSlide, created.id)).toEqual(
        customTriangleGeometry,
      );
    }
  });

  it('reads and whole-replaces live custom geometry with exact no-op semantics', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const shape = slide.addCustomShape(customTriangleGeometry, {
      name: 'Editable custom geometry',
      x: inches(2),
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
      line: { kind: 'line', color: { kind: 'srgb', value: '123456' }, width: 2 },
      arrows: { end: 'triangle' },
      shadow: { kind: 'outer' },
      hyperlink: { url: 'https://example.com/custom-edit' },
    });

    const first = shape.customGeometry;
    const second = shape.customGeometry;
    expect(first).toEqual(customTriangleGeometry);
    expect(second).toEqual(first);
    expect(second).not.toBe(first);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(first?.paths)).toBe(true);
    expect(Object.isFrozen(first?.paths[0])).toBe(true);
    expect(Object.isFrozen(first?.paths[0]?.commands)).toBe(true);
    expect(first?.paths[0]?.commands.every(Object.isFrozen)).toBe(true);
    expect(slide.getShapeCustomGeometry(shape.id)).toEqual(first);

    const noOp = packageSnapshot(pkg);
    shape.customGeometry = structuredClone(customTriangleGeometry);
    expect(packageSnapshot(pkg)).toEqual(noOp);
    expect(slide.shapes[0]).toBe(shape);

    for (const invalid of [
      undefined,
      null,
      { paths: [] },
      { paths: [{ width: 1, height: 1, commands: [{ kind: 'close' }] }] },
    ]) {
      expect(() => {
        shape.customGeometry = invalid as never;
      }).toThrow();
      expect(packageSnapshot(pkg)).toEqual(noOp);
    }

    const replacement: CustomGeometry = {
      paths: [
        {
          width: 200,
          height: 300,
          fill: 'none',
          stroke: false,
          commands: [
            { kind: 'moveTo', point: { x: 1, y: 2 } },
            {
              kind: 'cubicBezierTo',
              control1: { x: 3, y: 4 },
              control2: { x: 5, y: 6 },
              end: { x: 7, y: 8 },
            },
          ],
        },
        { width: 10, height: 20, extrusionOk: false, commands: [] },
      ],
    };
    const identity = shape;
    const transform = shape.transform;
    const fill = shape.fill;
    const line = shape.line;
    const arrows = shape.arrows;
    const shadow = shape.shadow;
    const hyperlink = shape.hyperlink;
    const relationships = slide.relationships;
    shape.customGeometry = replacement;
    expect(shape).toBe(identity);
    expect(slide.shapes[0]).toBe(identity);
    expect(shape.customGeometry).toEqual(replacement);
    expect(shape.transform).toEqual(transform);
    expect(shape.fill).toEqual(fill);
    expect(shape.line).toEqual(line);
    expect(shape.arrows).toEqual(arrows);
    expect(shape.shadow).toEqual(shadow);
    expect(shape.hyperlink).toEqual(hyperlink);
    expect(slide.relationships).toEqual(relationships);

    const part = pkg.requirePart(slide.partUri);
    pkg.setPart(
      slide.partUri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:gdLst/>',
        '<a:gdLst><a:gd name="x" fmla="unknown 1"/></a:gdLst>',
      ),
      part.contentType,
    );
    expect(shape.customGeometry).toBeUndefined();
    const unsupported = packageSnapshot(pkg);
    expect(() => {
      shape.customGeometry = customTriangleGeometry;
    }).toThrow(ModelParseError);
    expect(packageSnapshot(pkg)).toEqual(unsupported);
  });

  it('creates, freezes, edits, and rolls back live custom geometry guide formulas', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const mutable = structuredClone(customFormulaGeometry) as unknown as {
      adjustments: Array<{
        name: string;
        formula: { operator: string; operands: Array<number | string> };
      }>;
      guides: Array<{
        name: string;
        formula: { operator: string; operands: Array<number | string> };
      }>;
      paths: Array<{
        width: number;
        height: number;
        fill?: string;
        commands: Array<CustomGeometry['paths'][number]['commands'][number]>;
      }>;
    };
    const shape = slide.addCustomShape(mutable as unknown as CustomGeometry, {
      name: 'Formula geometry',
      x: inches(2),
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' } },
      line: { kind: 'line', color: { kind: 'srgb', value: '123ABC' }, width: 2 },
      arrows: { end: 'triangle' },
      shadow: { kind: 'outer' },
      hyperlink: { url: 'https://example.com/formula' },
    });
    const neighbor = slide.addShape('rect', { name: 'Formula neighbor' });
    const neighborBefore = neighbor.presetType;

    mutable.adjustments[0]!.name = 'changed';
    mutable.adjustments[0]!.formula.operands[0] = 1;
    mutable.guides.splice(0);
    mutable.paths[0]!.commands.splice(0);

    const first = shape.customGeometry;
    const second = shape.customGeometry;
    expect(first).toEqual(customFormulaGeometry);
    expect(second).toEqual(first);
    expect(second).not.toBe(first);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(first?.adjustments)).toBe(true);
    expect(Object.isFrozen(first?.adjustments?.[0])).toBe(true);
    expect(Object.isFrozen(first?.adjustments?.[0]?.formula)).toBe(true);
    expect(Object.isFrozen(first?.adjustments?.[0]?.formula.operands)).toBe(true);
    expect(Object.isFrozen(first?.guides)).toBe(true);
    expect(first?.guides?.every((guide) =>
      Object.isFrozen(guide)
      && Object.isFrozen(guide.formula)
      && Object.isFrozen(guide.formula.operands))).toBe(true);
    expect(readCreatedCustomGeometry(model, slide, shape.id)).toEqual(customFormulaGeometry);

    const noOp = packageSnapshot(pkg);
    shape.customGeometry = structuredClone(customFormulaGeometry);
    expect(packageSnapshot(pkg)).toEqual(noOp);

    const rollback = packageSnapshot(pkg);
    expect(() => pkg.transaction(() => {
      shape.customGeometry = customFormulaReplacement;
      expect(shape.customGeometry).toEqual(customFormulaReplacement);
      throw new Error('restore custom formula edit');
    })).toThrow('restore custom formula edit');
    expect(packageSnapshot(pkg)).toEqual(rollback);
    expect(shape.customGeometry).toEqual(customFormulaGeometry);

    const preserved = {
      name: shape.name,
      transform: shape.transform,
      fill: shape.fill,
      line: shape.line,
      arrows: shape.arrows,
      shadow: shape.shadow,
      hyperlink: shape.hyperlink,
      relationships: slide.relationships,
    };
    shape.customGeometry = customFormulaReplacement;
    expect(shape.customGeometry).toEqual(customFormulaReplacement);
    expect({
      name: shape.name,
      transform: shape.transform,
      fill: shape.fill,
      line: shape.line,
      arrows: shape.arrows,
      shadow: shape.shadow,
      hyperlink: shape.hyperlink,
      relationships: slide.relationships,
    }).toEqual(preserved);
    expect(slide.shapes[0]).toBe(shape);
    expect(slide.shapes[1]).toBe(neighbor);
    expect(neighbor.presetType).toBe(neighborBefore);
    const xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml).toContain('<a:gd name="adj1" fmla="val 50000"/>');
    expect(xml).toContain('<a:gd name="y1" fmla="min h x1"/>');
    expect(xml).toContain('<a:pt x="x1" y="y1"/>');
  });

  it('isolates, converts, and reopens custom geometry guide formulas in all six formats', async () => {
    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const pkg = await OpcPackage.open(await modelFixture(profile.presentationContentType));
      const model = new PresentationModel(pkg);
      const slide = model.addSlide();
      const shape = slide.addCustomShape(customFormulaGeometry, {
        name: `Formula ${profile.format}`,
        hyperlink: { url: `https://example.com/formula/${profile.format}` },
      });
      const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
      const duplicateShape = duplicate.shapes.find(({ id }) => id === shape.id) as ShapeModel;
      expect(duplicateShape.customGeometry).toEqual(customFormulaGeometry);

      shape.customGeometry = customFormulaReplacement;
      expect(shape.customGeometry).toEqual(customFormulaReplacement);
      expect(duplicateShape.customGeometry).toEqual(customFormulaGeometry);

      shape.presetType = 'diamond';
      expect(shape.customGeometry).toBeUndefined();
      expect(shape.presetType).toBe('diamond');
      shape.customGeometry = customFormulaReplacement;
      expect(shape.presetType).toBeUndefined();
      expect(shape.customGeometry).toEqual(customFormulaReplacement);
      expect(duplicateShape.customGeometry).toEqual(customFormulaGeometry);

      model.moveSlide(model.slides.indexOf(duplicate), 0);
      model.deleteSlide(model.slides.indexOf(duplicate));
      const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
      const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!;
      const reopenedShape = reopenedSlide.shapes.find(({ id }) => id === shape.id) as ShapeModel;
      expect(reopened.format).toBe(profile.format);
      expect(reopenedShape.name).toBe(`Formula ${profile.format}`);
      expect(reopenedShape.hyperlink).toEqual({
        url: `https://example.com/formula/${profile.format}`,
      });
      expect(reopenedShape.customGeometry).toEqual(customFormulaReplacement);
    }
  });

  it('keeps live custom geometry evaluation read-only and observes transform changes', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const shape = slide.addCustomShape(customEvaluationGeometry, {
      width: inches(2),
      height: inches(1),
    });
    const before = pkg.requirePart(slide.partUri).bytes.slice();
    const journal = [...pkg.mutations];

    const evaluated = shape.evaluateCustomGeometry();

    expect(evaluated).toMatchObject({
      context: { width: inches(2), height: inches(1) },
      adjustments: [{ name: 'adj', value: 25_000 }],
      guides: [
        { name: 'x1', value: inches(0.5) },
        { name: 'y1', value: inches(0.25) },
        { name: 'rad', value: inches(0.5) },
      ],
      handles: [{
        kind: 'xy',
        position: { x: inches(0.5), y: inches(0.25) },
        xGuide: 'adj',
        minX: 0,
        maxX: inches(2),
      }],
      connectionSites: [{
        angle: 5_400_000,
        position: { x: inches(0.5), y: inches(0.25) },
      }],
      textRectangle: {
        left: inches(0.5),
        top: inches(0.25),
        right: inches(2),
        bottom: inches(1),
      },
    });
    expect(evaluated?.paths[0]?.commands[1]).toEqual({
      kind: 'arcTo',
      widthRadius: inches(0.5),
      heightRadius: inches(0.5),
      startAngle: 0,
      sweepAngle: 5_400_000,
    });
    expect(Object.isFrozen(evaluated)).toBe(true);
    expect(Object.isFrozen(evaluated?.paths[0]?.commands[1])).toBe(true);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(before);
    expect(pkg.mutations).toEqual(journal);

    shape.setTransform({ width: inches(3), height: inches(1.5) });
    const resizedBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const resizedJournal = [...pkg.mutations];
    const resized = shape.evaluateCustomGeometry();
    expect(resized).not.toBe(evaluated);
    expect(resized?.context).toEqual({ width: inches(3), height: inches(1.5) });
    expect(resized?.guides).toEqual([
      { name: 'x1', value: inches(0.75) },
      { name: 'y1', value: inches(0.375) },
      { name: 'rad', value: inches(0.75) },
    ]);
    expect(resized?.textRectangle).toEqual({
      left: inches(0.75),
      top: inches(0.375),
      right: inches(3),
      bottom: inches(1.5),
    });
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(resizedBytes);
    expect(pkg.mutations).toEqual(resizedJournal);
  });

  it('keeps custom geometry evaluation failures read-only and returns undefined when absent', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const preset = model.addSlide().addShape('rect');
    expect(preset.evaluateCustomGeometry()).toBeUndefined();

    const malformedSlide = model.addSlide();
    const malformedShape = malformedSlide.addCustomShape(customEvaluationGeometry);
    const malformedPart = pkg.requirePart(malformedSlide.partUri);
    pkg.setPart(
      malformedSlide.partUri,
      new TextDecoder().decode(malformedPart.bytes).replace(
        'fmla="val 25000"',
        'fmla="unsupported 25000"',
      ),
      malformedPart.contentType,
    );
    const malformed = packageSnapshot(pkg);
    expect(malformedShape.customGeometry).toBeUndefined();
    expect(malformedShape.evaluateCustomGeometry()).toBeUndefined();
    expect(packageSnapshot(pkg)).toEqual(malformed);

    const unknownGeometry: CustomGeometry = {
      paths: [{
        width: 1,
        height: 1,
        commands: [{ kind: 'moveTo', point: { x: 'missing', y: 0 } }],
      }],
    };
    const unknownSlide = model.addSlide();
    const unknownShape = unknownSlide.addCustomShape(unknownGeometry);
    expect(unknownShape.customGeometry).toEqual(unknownGeometry);
    const before = packageSnapshot(pkg);
    try {
      unknownShape.evaluateCustomGeometry();
      throw new Error('Expected custom geometry evaluation to fail');
    } catch (error) {
      expect(error).toBeInstanceOf(CustomGeometryEvaluationError);
      expect(error).toMatchObject({ code: 'unknown-token', token: 'missing' });
    }
    expect(packageSnapshot(pkg)).toEqual(before);
  });

  it('preserves custom geometry evaluation through duplicate, move, and all formats', async () => {
    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const pkg = await OpcPackage.open(await modelFixture(profile.presentationContentType));
      const model = new PresentationModel(pkg);
      const slide = model.addSlide();
      const shape = slide.addCustomShape(customEvaluationGeometry, {
        name: `Evaluator ${profile.format}`,
        width: inches(2),
        height: inches(1),
      });
      const expected = shape.evaluateCustomGeometry();
      const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
      const duplicateShape = duplicate.shapes.find(({ id }) => id === shape.id) as ShapeModel;
      model.moveSlide(model.slides.indexOf(duplicate), 0);
      expect(duplicateShape.evaluateCustomGeometry()).toEqual(expected);

      const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
      for (const partUri of [slide.partUri, duplicate.partUri]) {
        const reopenedSlide = reopened.slides.find((candidate) => candidate.partUri === partUri)!;
        const reopenedShape = reopenedSlide.shapes.find(({ id }) => id === shape.id) as ShapeModel;
        expect(reopenedShape.name).toBe(`Evaluator ${profile.format}`);
        expect(reopenedShape.evaluateCustomGeometry()).toEqual(expected);
      }
    }
  });

  it('creates, freezes, edits, and rolls back live custom geometry adjustment handles', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const mutable = structuredClone(customHandleGeometry) as unknown as {
      handles: [
        { position: { x: string | number; y: string | number }; minX: number },
        { position: { x: string | number; y: string | number }; maxAngle: string | number },
      ];
    };
    let shape = slide.addCustomShape(mutable as unknown as CustomGeometry, {
      name: 'Handle geometry',
      x: inches(2),
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' } },
      line: { kind: 'line', color: { kind: 'srgb', value: '123ABC' }, width: 2 },
      arrows: { end: 'triangle' },
      shadow: { kind: 'outer' },
      hyperlink: { url: 'https://example.com/handles' },
    });
    const createdPart = pkg.requirePart(slide.partUri);
    pkg.setPart(
      slide.partUri,
      new TextDecoder().decode(createdPart.bytes).replace(
        '</p:spPr></p:sp>',
        '<a:extLst><a:ext uri="urn:handles"><x:keep xmlns:x="urn:handles">' +
        'EXTENSION</x:keep></a:ext></a:extLst></p:spPr>' +
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>HANDLE TEXT</a:t>' +
        '</a:r></a:p></p:txBody></p:sp>',
      ),
      createdPart.contentType,
    );
    const shapeId = shape.id;
    shape = slide.shapes.find(({ id }) => id === shapeId) as ShapeModel;
    const neighbor = slide.addShape('rect', { name: 'Handle neighbor' });

    mutable.handles[0].position.x = 'changed';
    mutable.handles[0].minX = -1;
    mutable.handles[1].maxAngle = 1;
    mutable.handles.reverse();

    const first = shape.customGeometry;
    const second = shape.customGeometry;
    expect(first).toEqual(customHandleGeometry);
    expect(second).toEqual(first);
    expect(second).not.toBe(first);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(first?.handles)).toBe(true);
    expect(first?.handles?.every((handle) =>
      Object.isFrozen(handle) && Object.isFrozen(handle.position))).toBe(true);
    expect(readCreatedCustomGeometry(model, slide, shape.id)).toEqual(customHandleGeometry);

    const noOp = packageSnapshot(pkg);
    shape.customGeometry = structuredClone(customHandleGeometry);
    expect(packageSnapshot(pkg)).toEqual(noOp);

    const rollback = packageSnapshot(pkg);
    expect(() => pkg.transaction(() => {
      shape.customGeometry = customHandleReplacement;
      expect(shape.customGeometry).toEqual(customHandleReplacement);
      throw new Error('restore custom handle edit');
    })).toThrow('restore custom handle edit');
    expect(packageSnapshot(pkg)).toEqual(rollback);
    expect(shape.customGeometry).toEqual(customHandleGeometry);

    const identity = shape;
    const preserved = {
      name: shape.name,
      transform: shape.transform,
      fill: shape.fill,
      line: shape.line,
      arrows: shape.arrows,
      shadow: shape.shadow,
      hyperlink: shape.hyperlink,
      relationships: slide.relationships,
    };
    const beforeEdit = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    shape.customGeometry = customHandleReplacement;
    const afterEdit = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    const withoutGeometry = (source: string) => source.replace(
      /<a:custGeom>[\s\S]*?<\/a:custGeom>/,
      '<a:custGeom/>',
    );
    expect(shape).toBe(identity);
    expect(slide.shapes[0]).toBe(identity);
    expect(slide.shapes[1]).toBe(neighbor);
    expect(shape.customGeometry).toEqual(customHandleReplacement);
    expect({
      name: shape.name,
      transform: shape.transform,
      fill: shape.fill,
      line: shape.line,
      arrows: shape.arrows,
      shadow: shape.shadow,
      hyperlink: shape.hyperlink,
      relationships: slide.relationships,
    }).toEqual(preserved);
    expect(withoutGeometry(afterEdit)).toBe(withoutGeometry(beforeEdit));
    expect(afterEdit).toContain('<a:effectLst>');
    expect(afterEdit).toContain('EXTENSION');
    expect(afterEdit).toContain('<a:t>HANDLE TEXT</a:t>');
    expect(afterEdit.indexOf('<a:ahPolar')).toBeLessThan(afterEdit.indexOf('<a:ahXY'));
  });

  it('isolates, converts, and reopens custom geometry adjustment handles in all six formats', async () => {
    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const pkg = await OpcPackage.open(await modelFixture(profile.presentationContentType));
      const model = new PresentationModel(pkg);
      const slide = model.addSlide();
      const beforeCreation = packageSnapshot(pkg);
      let rolledBack: ShapeModel | undefined;
      expect(() => pkg.transaction(() => {
        rolledBack = slide.addCustomShape(customHandleGeometry, {
          hyperlink: { url: `https://example.com/handle-rollback/${profile.format}` },
        });
        throw new Error('restore custom handle creation');
      })).toThrow('restore custom handle creation');
      expect(packageSnapshot(pkg)).toEqual(beforeCreation);
      expect(() => rolledBack!.name).toThrow(ModelParseError);

      const shape = slide.addCustomShape(customHandleGeometry, {
        name: `Handles ${profile.format}`,
        hyperlink: { url: `https://example.com/handles/${profile.format}` },
      });
      expect(shape.id).toBe(2);
      const beforeEdit = packageSnapshot(pkg);
      expect(() => pkg.transaction(() => {
        shape.customGeometry = customHandleReplacement;
        throw new Error('restore custom handle replacement');
      })).toThrow('restore custom handle replacement');
      expect(packageSnapshot(pkg)).toEqual(beforeEdit);
      expect(shape.customGeometry).toEqual(customHandleGeometry);

      const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
      const duplicateShape = duplicate.shapes.find(({ id }) => id === shape.id) as ShapeModel;
      shape.customGeometry = customHandleReplacement;
      expect(shape.customGeometry).toEqual(customHandleReplacement);
      expect(duplicateShape.customGeometry).toEqual(customHandleGeometry);

      const convertible = slide.addShape('diamond', {
        name: `Handle conversion ${profile.format}`,
        fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
        hyperlink: { url: `https://example.com/handle-conversion/${profile.format}` },
      });
      const convertibleIdentity = convertible;
      convertible.customGeometry = customHandleGeometry;
      expect(convertible).toBe(convertibleIdentity);
      expect(convertible.customGeometry).toEqual(customHandleGeometry);
      convertible.presetType = 'ellipse';
      expect(convertible).toBe(convertibleIdentity);
      expect(convertible.presetType).toBe('ellipse');
      expect(convertible.customGeometry).toBeUndefined();
      expect(convertible.name).toBe(`Handle conversion ${profile.format}`);
      expect(convertible.fill).toEqual({
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
      });
      expect(convertible.hyperlink).toEqual({
        url: `https://example.com/handle-conversion/${profile.format}`,
      });

      model.moveSlide(model.slides.indexOf(duplicate), 0);
      model.deleteSlide(model.slides.indexOf(duplicate));
      const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
      const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!;
      const reopenedShape = reopenedSlide.shapes.find(({ id }) => id === shape.id) as ShapeModel;
      const reopenedConvertible = reopenedSlide.shapes.find(
        ({ id }) => id === convertible.id,
      ) as ShapeModel;
      expect(reopened.format).toBe(profile.format);
      expect(reopenedShape.name).toBe(`Handles ${profile.format}`);
      expect(reopenedShape.hyperlink).toEqual({
        url: `https://example.com/handles/${profile.format}`,
      });
      expect(reopenedShape.customGeometry).toEqual(customHandleReplacement);
      expect(reopenedConvertible.presetType).toBe('ellipse');
      expect(reopenedConvertible.customGeometry).toBeUndefined();
      expect(reopenedConvertible.name).toBe(`Handle conversion ${profile.format}`);
    }
  });

  it('creates, freezes, edits, and rolls back live custom geometry connection sites', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const mutable = structuredClone(customConnectionGeometry) as unknown as {
      connectionSites: Array<{
        angle: string | number;
        position: { x: string | number; y: string | number };
      }>;
    };
    let shape = slide.addCustomShape(mutable as unknown as CustomGeometry, {
      name: 'Connection geometry',
      x: inches(2),
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' } },
      line: { kind: 'line', color: { kind: 'srgb', value: '123ABC' }, width: 2 },
      arrows: { end: 'triangle' },
      shadow: { kind: 'outer' },
      hyperlink: { url: 'https://example.com/connections' },
    });
    const createdPart = pkg.requirePart(slide.partUri);
    pkg.setPart(
      slide.partUri,
      new TextDecoder().decode(createdPart.bytes).replace(
        '</p:spPr></p:sp>',
        '<a:extLst><a:ext uri="urn:connections"><x:keep xmlns:x="urn:connections">' +
        'EXTENSION</x:keep></a:ext></a:extLst></p:spPr>' +
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>CONNECTION TEXT</a:t>' +
        '</a:r></a:p></p:txBody></p:sp>',
      ),
      createdPart.contentType,
    );
    const shapeId = shape.id;
    shape = slide.shapes.find(({ id }) => id === shapeId) as ShapeModel;
    const neighbor = slide.addShape('rect', { name: 'Connection neighbor' });

    mutable.connectionSites[0]!.angle = 1;
    mutable.connectionSites[0]!.position.x = 'changed';
    mutable.connectionSites.reverse();

    const first = shape.customGeometry;
    const second = shape.customGeometry;
    expect(first).toEqual(customConnectionGeometry);
    expect(second).toEqual(first);
    expect(second).not.toBe(first);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(first?.connectionSites)).toBe(true);
    expect(first?.connectionSites?.every((site) =>
      Object.isFrozen(site) && Object.isFrozen(site.position))).toBe(true);
    expect(readCreatedCustomGeometry(model, slide, shape.id)).toEqual(customConnectionGeometry);

    const noOp = packageSnapshot(pkg);
    shape.customGeometry = structuredClone(customConnectionGeometry);
    expect(packageSnapshot(pkg)).toEqual(noOp);

    const rollback = packageSnapshot(pkg);
    expect(() => pkg.transaction(() => {
      shape.customGeometry = customConnectionReplacement;
      expect(shape.customGeometry).toEqual(customConnectionReplacement);
      throw new Error('restore custom connection edit');
    })).toThrow('restore custom connection edit');
    expect(packageSnapshot(pkg)).toEqual(rollback);
    expect(shape.customGeometry).toEqual(customConnectionGeometry);

    const identity = shape;
    const preserved = {
      name: shape.name,
      transform: shape.transform,
      fill: shape.fill,
      line: shape.line,
      arrows: shape.arrows,
      shadow: shape.shadow,
      hyperlink: shape.hyperlink,
      relationships: slide.relationships,
    };
    const beforeEdit = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    shape.customGeometry = customConnectionReplacement;
    const afterEdit = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    const withoutGeometry = (source: string) => source.replace(
      /<a:custGeom>[\s\S]*?<\/a:custGeom>/,
      '<a:custGeom/>',
    );
    expect(shape).toBe(identity);
    expect(slide.shapes[0]).toBe(identity);
    expect(slide.shapes[1]).toBe(neighbor);
    expect(shape.customGeometry).toEqual(customConnectionReplacement);
    expect({
      name: shape.name,
      transform: shape.transform,
      fill: shape.fill,
      line: shape.line,
      arrows: shape.arrows,
      shadow: shape.shadow,
      hyperlink: shape.hyperlink,
      relationships: slide.relationships,
    }).toEqual(preserved);
    expect(withoutGeometry(afterEdit)).toBe(withoutGeometry(beforeEdit));
    expect(afterEdit).toContain('<a:effectLst>');
    expect(afterEdit).toContain('EXTENSION');
    expect(afterEdit).toContain('<a:t>CONNECTION TEXT</a:t>');
    expect(afterEdit.indexOf('<a:cxn ang="-5400000"')).toBeLessThan(
      afterEdit.indexOf('<a:cxn ang="adjAng"'),
    );
  });

  it('isolates, converts, and reopens custom geometry connection sites in all six formats', async () => {
    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const pkg = await OpcPackage.open(await modelFixture(profile.presentationContentType));
      const model = new PresentationModel(pkg);
      const slide = model.addSlide();
      const beforeCreation = packageSnapshot(pkg);
      let rolledBack: ShapeModel | undefined;
      expect(() => pkg.transaction(() => {
        rolledBack = slide.addCustomShape(customConnectionGeometry, {
          hyperlink: { url: `https://example.com/connection-rollback/${profile.format}` },
        });
        throw new Error('restore custom connection creation');
      })).toThrow('restore custom connection creation');
      expect(packageSnapshot(pkg)).toEqual(beforeCreation);
      expect(() => rolledBack!.name).toThrow(ModelParseError);

      const shape = slide.addCustomShape(customConnectionGeometry, {
        name: `Connections ${profile.format}`,
        fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
        hyperlink: { url: `https://example.com/connections/${profile.format}` },
      });
      const beforeEdit = packageSnapshot(pkg);
      expect(() => pkg.transaction(() => {
        shape.customGeometry = customConnectionReplacement;
        throw new Error('restore custom connection replacement');
      })).toThrow('restore custom connection replacement');
      expect(packageSnapshot(pkg)).toEqual(beforeEdit);
      expect(shape.customGeometry).toEqual(customConnectionGeometry);

      const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
      const duplicateShape = duplicate.shapes.find(({ id }) => id === shape.id) as ShapeModel;
      duplicateShape.customGeometry = customConnectionReplacement;
      expect(duplicateShape.customGeometry).toEqual(customConnectionReplacement);
      expect(shape.customGeometry).toEqual(customConnectionGeometry);

      const identity = shape;
      shape.presetType = 'diamond';
      expect(shape).toBe(identity);
      expect(shape.presetType).toBe('diamond');
      expect(shape.customGeometry).toBeUndefined();
      shape.customGeometry = customConnectionReplacement;
      expect(shape).toBe(identity);
      expect(shape.presetType).toBeUndefined();
      expect(shape.customGeometry).toEqual(customConnectionReplacement);
      expect(shape.name).toBe(`Connections ${profile.format}`);
      expect(shape.fill).toEqual({
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
      });
      expect(shape.hyperlink).toEqual({
        url: `https://example.com/connections/${profile.format}`,
      });

      model.moveSlide(model.slides.indexOf(duplicate), 0);
      model.deleteSlide(model.slides.indexOf(duplicate));
      const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
      const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!;
      const reopenedShape = reopenedSlide.shapes.find(({ id }) => id === shape.id) as ShapeModel;
      expect(reopened.format).toBe(profile.format);
      expect(reopenedShape.name).toBe(`Connections ${profile.format}`);
      expect(reopenedShape.fill).toEqual({
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
      });
      expect(reopenedShape.hyperlink).toEqual({
        url: `https://example.com/connections/${profile.format}`,
      });
      expect(reopenedShape.customGeometry).toEqual(customConnectionReplacement);
    }
  });

  it('creates, freezes, edits, resets, and rolls back live custom geometry text rectangles', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const mutable = structuredClone(customTextRectangleGeometry) as unknown as {
      textRectangle: {
        left: string | number;
        top: string | number;
        right: string | number;
        bottom: string | number;
      };
    };
    let shape = slide.addCustomShape(mutable as unknown as CustomGeometry, {
      name: 'Text rectangle geometry',
      x: inches(2),
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' } },
      line: { kind: 'line', color: { kind: 'srgb', value: '123ABC' }, width: 2 },
      arrows: { end: 'triangle' },
      shadow: { kind: 'outer' },
      hyperlink: { url: 'https://example.com/text-rectangle' },
    });
    const createdPart = pkg.requirePart(slide.partUri);
    pkg.setPart(
      slide.partUri,
      new TextDecoder().decode(createdPart.bytes).replace(
        '</p:spPr></p:sp>',
        '<a:extLst><a:ext uri="urn:text-rectangle">' +
        '<x:keep xmlns:x="urn:text-rectangle">EXTENSION</x:keep></a:ext>' +
        '</a:extLst></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p>' +
        '<a:r><a:t>TEXT RECTANGLE TEXT</a:t></a:r></a:p></p:txBody></p:sp>',
      ),
      createdPart.contentType,
    );
    const shapeId = shape.id;
    shape = slide.shapes.find(({ id }) => id === shapeId) as ShapeModel;
    const neighbor = slide.addShape('rect', { name: 'Text rectangle neighbor' });

    mutable.textRectangle.left = 'changed';
    mutable.textRectangle.top = 1;
    mutable.textRectangle.right = 2;
    mutable.textRectangle.bottom = 3;

    const first = shape.customGeometry;
    const second = shape.customGeometry;
    expect(first).toEqual(customTextRectangleGeometry);
    expect(second).toEqual(first);
    expect(second).not.toBe(first);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(first?.textRectangle)).toBe(true);
    expect(readCreatedCustomGeometry(model, slide, shape.id)).toEqual(
      customTextRectangleGeometry,
    );

    const noOp = packageSnapshot(pkg);
    shape.customGeometry = structuredClone(customTextRectangleGeometry);
    expect(packageSnapshot(pkg)).toEqual(noOp);

    const rollback = packageSnapshot(pkg);
    expect(() => pkg.transaction(() => {
      shape.customGeometry = customTextRectangleReplacement;
      expect(shape.customGeometry).toEqual(customTextRectangleReplacement);
      throw new Error('restore custom text rectangle edit');
    })).toThrow('restore custom text rectangle edit');
    expect(packageSnapshot(pkg)).toEqual(rollback);
    expect(shape.customGeometry).toEqual(customTextRectangleGeometry);

    const identity = shape;
    const preserved = {
      name: shape.name,
      transform: shape.transform,
      fill: shape.fill,
      line: shape.line,
      arrows: shape.arrows,
      shadow: shape.shadow,
      hyperlink: shape.hyperlink,
      relationships: slide.relationships,
    };
    const beforeEdit = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    shape.customGeometry = customTextRectangleReplacement;
    const afterEdit = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    const withoutGeometry = (source: string) => source.replace(
      /<a:custGeom>[\s\S]*?<\/a:custGeom>/,
      '<a:custGeom/>',
    );
    expect(shape).toBe(identity);
    expect(slide.shapes[0]).toBe(identity);
    expect(slide.shapes[1]).toBe(neighbor);
    expect(shape.customGeometry).toEqual(customTextRectangleReplacement);
    expect({
      name: shape.name,
      transform: shape.transform,
      fill: shape.fill,
      line: shape.line,
      arrows: shape.arrows,
      shadow: shape.shadow,
      hyperlink: shape.hyperlink,
      relationships: slide.relationships,
    }).toEqual(preserved);
    expect(withoutGeometry(afterEdit)).toBe(withoutGeometry(beforeEdit));
    expect(afterEdit).toContain('<a:effectLst>');
    expect(afterEdit).toContain('EXTENSION');
    expect(afterEdit).toContain('<a:t>TEXT RECTANGLE TEXT</a:t>');
    expect(afterEdit).toContain('<a:rect l="0" t="t" r="80000" b="b"/>');

    const { textRectangle: _textRectangle, ...defaultTextRectangleGeometry } =
      customTextRectangleReplacement;
    const beforeReset = packageSnapshot(pkg);
    expect(() => pkg.transaction(() => {
      shape.customGeometry = defaultTextRectangleGeometry;
      expect(Object.hasOwn(shape.customGeometry!, 'textRectangle')).toBe(false);
      throw new Error('restore custom text rectangle reset');
    })).toThrow('restore custom text rectangle reset');
    expect(packageSnapshot(pkg)).toEqual(beforeReset);
    expect(shape.customGeometry).toEqual(customTextRectangleReplacement);

    shape.customGeometry = defaultTextRectangleGeometry;
    expect(Object.hasOwn(shape.customGeometry!, 'textRectangle')).toBe(false);
    expect(new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes))
      .toContain('<a:rect l="l" t="t" r="r" b="b"/>');
  });

  it('isolates, converts, and reopens custom geometry text rectangles in all six formats', async () => {
    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const pkg = await OpcPackage.open(await modelFixture(profile.presentationContentType));
      const model = new PresentationModel(pkg);
      const slide = model.addSlide();
      const beforeCreation = packageSnapshot(pkg);
      let rolledBack: ShapeModel | undefined;
      expect(() => pkg.transaction(() => {
        rolledBack = slide.addCustomShape(customTextRectangleGeometry, {
          hyperlink: { url: `https://example.com/text-rectangle-rollback/${profile.format}` },
        });
        throw new Error('restore custom text rectangle creation');
      })).toThrow('restore custom text rectangle creation');
      expect(packageSnapshot(pkg)).toEqual(beforeCreation);
      expect(() => rolledBack!.name).toThrow(ModelParseError);

      const shape = slide.addCustomShape(customTextRectangleGeometry, {
        name: `Text rectangle ${profile.format}`,
        fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
        hyperlink: { url: `https://example.com/text-rectangles/${profile.format}` },
      });
      expect(shape.id).toBe(2);
      const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
      const duplicateShape = duplicate.shapes.find(({ id }) => id === shape.id) as ShapeModel;
      duplicateShape.customGeometry = customTextRectangleReplacement;
      expect(duplicateShape.customGeometry).toEqual(customTextRectangleReplacement);
      expect(shape.customGeometry).toEqual(customTextRectangleGeometry);

      const identity = shape;
      shape.presetType = 'diamond';
      expect(shape).toBe(identity);
      expect(shape.presetType).toBe('diamond');
      expect(shape.customGeometry).toBeUndefined();
      shape.customGeometry = customTextRectangleReplacement;
      expect(shape).toBe(identity);
      expect(shape.presetType).toBeUndefined();
      expect(shape.customGeometry).toEqual(customTextRectangleReplacement);

      const { textRectangle: _textRectangle, ...defaultTextRectangleGeometry } =
        customTextRectangleReplacement;
      const beforeReset = packageSnapshot(pkg);
      expect(() => pkg.transaction(() => {
        shape.customGeometry = defaultTextRectangleGeometry;
        expect(Object.hasOwn(shape.customGeometry!, 'textRectangle')).toBe(false);
        throw new Error('restore custom text rectangle reset');
      })).toThrow('restore custom text rectangle reset');
      expect(packageSnapshot(pkg)).toEqual(beforeReset);
      expect(shape.customGeometry).toEqual(customTextRectangleReplacement);
      expect(shape.name).toBe(`Text rectangle ${profile.format}`);
      expect(shape.fill).toEqual({
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
      });
      expect(shape.hyperlink).toEqual({
        url: `https://example.com/text-rectangles/${profile.format}`,
      });

      model.moveSlide(model.slides.indexOf(duplicate), 0);
      model.deleteSlide(model.slides.indexOf(duplicate));
      const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
      const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!;
      const reopenedShape = reopenedSlide.shapes.find(({ id }) => id === shape.id) as ShapeModel;
      expect(reopened.format).toBe(profile.format);
      expect(reopenedShape.name).toBe(`Text rectangle ${profile.format}`);
      expect(reopenedShape.fill).toEqual({
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
      });
      expect(reopenedShape.hyperlink).toEqual({
        url: `https://example.com/text-rectangles/${profile.format}`,
      });
      expect(reopenedShape.customGeometry).toEqual(customTextRectangleReplacement);
    }
  });

  it('reads and whole-replaces custom connections and text rectangles', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const connectionSlide = model.addSlide();
    const connectionShape = connectionSlide.addCustomShape(customHandleGeometry, {
      name: 'connections',
    });
    const connectionPart = pkg.requirePart(connectionSlide.partUri);
    const connectionSource = new TextDecoder().decode(connectionPart.bytes);
    expect(connectionSource).toContain('<a:ahXY');
    expect(connectionSource).toContain('<a:cxnLst/>');
    pkg.setPart(
      connectionSlide.partUri,
      connectionSource.replace(
        '<a:cxnLst/>',
        '<a:cxnLst><a:cxn ang="0"><a:pos x="0" y="0"/></a:cxn></a:cxnLst>',
      ),
      connectionPart.contentType,
    );
    expect(connectionShape.customGeometry).toEqual({
      ...customHandleGeometry,
      connectionSites: [{ angle: 0, position: { x: 0, y: 0 } }],
    });
    connectionShape.customGeometry = customConnectionReplacement;
    expect(connectionShape.customGeometry).toEqual(customConnectionReplacement);

    const rectangleSlide = model.addSlide();
    const rectangleShape = rectangleSlide.addCustomShape(customHandleGeometry, {
      name: 'text rectangle',
    });
    const rectanglePart = pkg.requirePart(rectangleSlide.partUri);
    const rectangleSource = new TextDecoder().decode(rectanglePart.bytes);
    expect(rectangleSource).toContain('<a:rect l="l" t="t" r="r" b="b"/>');
    pkg.setPart(
      rectangleSlide.partUri,
      rectangleSource.replace(
        '<a:rect l="l" t="t" r="r" b="b"/>',
        '<a:rect l="0" t="t" r="r" b="b"/>',
      ),
      rectanglePart.contentType,
    );
    expect(rectangleShape.customGeometry).toEqual({
      ...customHandleGeometry,
      textRectangle: { left: 0, top: 't', right: 'r', bottom: 'b' },
    });
    const rectangleReplacement: CustomGeometry = {
      ...customConnectionReplacement,
      textRectangle: { left: 'x1', top: 10_000, right: 'r', bottom: 90_000 },
    };
    rectangleShape.customGeometry = rectangleReplacement;
    expect(rectangleShape.customGeometry).toEqual(rectangleReplacement);
    const beforeNoOp = packageSnapshot(pkg);
    rectangleShape.customGeometry = structuredClone(rectangleReplacement);
    expect(packageSnapshot(pkg)).toEqual(beforeNoOp);

    const { textRectangle: _textRectangle, ...resetRectangle } = rectangleReplacement;
    rectangleShape.customGeometry = resetRectangle;
    expect(rectangleShape.customGeometry).toEqual(resetRectangle);
    expect(Object.hasOwn(rectangleShape.customGeometry!, 'textRectangle')).toBe(false);
    expect(new TextDecoder().decode(pkg.requirePart(rectangleSlide.partUri).bytes))
      .toContain('<a:rect l="l" t="t" r="r" b="b"/>');
  });

  it('converts preset and custom geometry without changing live identity or unrelated state', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const shape = slide.addShape('blockArc', {
      name: 'Convertible geometry',
      adjustments: [
        { name: 'adj1', value: 16_200_000 },
        { name: 'adj2', value: 0 },
      ],
      x: inches(2),
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' } },
      line: { kind: 'line', color: { kind: 'srgb', value: '123ABC' } },
      arrows: { end: 'triangle' },
      shadow: { kind: 'outer' },
      hyperlink: { url: 'https://example.com/convert' },
    });
    const text = slide.addText('Keep text', { name: 'Convertible text' });
    const identity = shape;
    const preserved = {
      name: shape.name,
      transform: shape.transform,
      fill: shape.fill,
      line: shape.line,
      arrows: shape.arrows,
      shadow: shape.shadow,
      hyperlink: shape.hyperlink,
    };

    shape.customGeometry = customTriangleGeometry;
    expect(shape).toBe(identity);
    expect(slide.shapes[0]).toBe(identity);
    expect(shape.presetType).toBeUndefined();
    expect(shape.adjustments).toBeUndefined();
    expect(shape.customGeometry).toEqual(customTriangleGeometry);
    expect({
      name: shape.name,
      transform: shape.transform,
      fill: shape.fill,
      line: shape.line,
      arrows: shape.arrows,
      shadow: shape.shadow,
      hyperlink: shape.hyperlink,
    }).toEqual(preserved);

    text.customGeometry = customTriangleGeometry;
    expect(text.text).toBe('Keep text');
    expect(text.customGeometry).toEqual(customTriangleGeometry);
    text.presetType = 'ellipse';
    expect(text.text).toBe('Keep text');
    expect(text.customGeometry).toBeUndefined();
    expect(text.presetType).toBe('ellipse');

    const customSnapshot = packageSnapshot(pkg);
    expect(() => pkg.transaction(() => {
      shape.presetType = 'star5';
      expect(shape.customGeometry).toBeUndefined();
      throw new Error('restore custom conversion');
    })).toThrow('restore custom conversion');
    expect(packageSnapshot(pkg)).toEqual(customSnapshot);
    expect(shape.customGeometry).toEqual(customTriangleGeometry);

    const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
    const duplicateShape = duplicate.shapes.find(({ id }) => id === shape.id) as ShapeModel;
    shape.presetType = 'ellipse';
    expect(shape.presetType).toBe('ellipse');
    expect(shape.adjustments).toEqual([]);
    expect(shape.customGeometry).toBeUndefined();
    expect(duplicateShape.customGeometry).toEqual(customTriangleGeometry);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!;
    const reopenedShape = reopenedSlide.shapes.find(({ id }) => id === shape.id) as ShapeModel;
    expect(reopenedShape.presetType).toBe('ellipse');
    expect(reopenedShape.customGeometry).toBeUndefined();
    expect(reopenedShape.name).toBe('Convertible geometry');
    expect(reopenedShape.hyperlink).toEqual({ url: 'https://example.com/convert' });

    const malformedSlide = model.addSlide();
    const malformedShape = malformedSlide.addShape('rect');
    const part = pkg.requirePart(malformedSlide.partUri);
    pkg.setPart(
      malformedSlide.partUri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:custGeom/>',
      ),
      part.contentType,
    );
    const malformed = packageSnapshot(pkg);
    expect(() => {
      malformedShape.customGeometry = customTriangleGeometry;
    }).toThrow(ModelParseError);
    expect(packageSnapshot(pkg)).toEqual(malformed);
  });

  it('creates preset shape adjustments with detached ordered values and unchanged empty bytes', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const omittedSlide = model.addSlide();
    const undefinedSlide = model.addSlide();
    const emptySlide = model.addSlide();
    omittedSlide.addShape('rect');
    undefinedSlide.addShape('rect', { adjustments: undefined } as never);
    emptySlide.addShape('rect', { adjustments: [] });
    expect(pkg.requirePart(undefinedSlide.partUri).bytes).toEqual(
      pkg.requirePart(omittedSlide.partUri).bytes,
    );
    expect(pkg.requirePart(emptySlide.partUri).bytes).toEqual(
      pkg.requirePart(omittedSlide.partUri).bytes,
    );

    const slide = model.addSlide();
    const relationships = slide.relationships;
    const adjustments: { name: string; value: number }[] = [
      { name: 'adj1', value: 16_200_000 },
      { name: 'adj2', value: 0 },
      { name: 'adj3', value: 25_000 },
    ];
    const blockArc = slide.addShape('blockArc', {
      name: 'Adjusted block arc',
      adjustments,
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' } },
      line: { kind: 'line', color: { kind: 'srgb', value: '123ABC' } },
      arrows: { end: 'triangle' },
      shadow: { kind: 'outer' },
      x: inches(2),
    });
    adjustments[0]!.value = 1;

    expect(slide.shapes).toEqual([blockArc]);
    expect(slide.shapes[0]).toBe(blockArc);
    expect(slide.relationships).toEqual(relationships);
    expect(readCreatedShapeAdjustments(model, slide, blockArc.id)).toEqual([
      { name: 'adj1', value: 16_200_000 },
      { name: 'adj2', value: 0 },
      { name: 'adj3', value: 25_000 },
    ]);
    const xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml).toContain(
      '<a:prstGeom prst="blockArc"><a:avLst>' +
      '<a:gd name="adj1" fmla="val 16200000"/>' +
      '<a:gd name="adj2" fmla="val 0"/>' +
      '<a:gd name="adj3" fmla="val 25000"/></a:avLst></a:prstGeom>',
    );

    const publicValue: readonly ShapeAdjustment[] = readCreatedShapeAdjustments(
      model,
      slide,
      blockArc.id,
    )!;
    expect(publicValue).toHaveLength(3);
  });

  it('rejects invalid preset shape adjustment creation without package, identity, or ID mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const before = packageSnapshot(pkg);
    const shapes = slide.shapes;
    let calls = 0;
    const optionsGetter = Object.defineProperty({}, 'adjustments', {
      enumerable: true,
      get() {
        calls += 1;
        return [];
      },
    });
    const entryGetter = Object.defineProperty({ name: 'adj' }, 'value', {
      enumerable: true,
      get() {
        calls += 1;
        return 1;
      },
    });
    const sparse = new Array(2);
    sparse[1] = { name: 'adj', value: 1 };
    for (const adjustments of [
      null,
      {},
      sparse,
      [entryGetter],
      [{ name: '', value: 1 }],
      [{ name: 'adj', value: 1.5 }],
      [{ name: 'adj', value: 1 }, { name: 'adj', value: 2 }],
      [{ name: 'adj', value: 1, extra: true }],
    ]) {
      expect(() => slide.addShape('rect', { adjustments } as never)).toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
      expect(slide.shapes).toEqual(shapes);
    }
    expect(() => slide.addShape('rect', optionsGetter as never)).toThrow(/data property/);
    expect(calls).toBe(0);
    expect(packageSnapshot(pkg)).toEqual(before);
    expect(slide.addShape('rect').id).toBe(2);
  });

  it('rolls back and reopens preset shape adjustment creation in all six formats', async () => {
    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const pkg = await OpcPackage.open(await modelFixture(profile.presentationContentType));
      const model = new PresentationModel(pkg);
      const slide = model.addSlide();
      const before = packageSnapshot(pkg);
      let rolledBack: ShapeModel | undefined;
      expect(() => pkg.transaction(() => {
        rolledBack = slide.addShape('pie', {
          adjustments: [
            { name: 'adj1', value: 16_200_000 },
            { name: 'adj2', value: 0 },
          ],
        });
        throw new Error('restore adjustment creation');
      })).toThrow('restore adjustment creation');
      expect(packageSnapshot(pkg)).toEqual(before);
      expect(() => rolledBack!.name).toThrow(ModelParseError);

      const created = slide.addShape('blockArc', {
        adjustments: [
          { name: 'adj1', value: 16_200_000 },
          { name: 'adj2', value: 0 },
          { name: 'adj3', value: 25_000 },
        ],
      });
      const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
      expect(readCreatedShapeAdjustments(model, duplicate, created.id)).toEqual([
        { name: 'adj1', value: 16_200_000 },
        { name: 'adj2', value: 0 },
        { name: 'adj3', value: 25_000 },
      ]);

      const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
      const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!;
      expect(reopened.format).toBe(profile.format);
      expect(readCreatedShapeAdjustments(reopened, reopenedSlide, created.id)).toEqual([
        { name: 'adj1', value: 16_200_000 },
        { name: 'adj2', value: 0 },
        { name: 'adj3', value: 25_000 },
      ]);
    }
  });

  it('reads, replaces, clears, and rolls back live preset shape adjustments exactly', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const shape = slide.addShape('blockArc', {
      adjustments: [
        { name: 'adj1', value: 16_200_000 },
        { name: 'adj2', value: 0 },
        { name: 'adj3', value: 25_000 },
      ],
      fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent2' } },
      line: { kind: 'line', color: { kind: 'srgb', value: '123ABC' } },
      shadow: { kind: 'outer' },
    });

    const initial = shape.adjustments;
    expect(initial).toEqual([
      { name: 'adj1', value: 16_200_000 },
      { name: 'adj2', value: 0 },
      { name: 'adj3', value: 25_000 },
    ]);
    expect(slide.getShapeAdjustments(shape.id)).toEqual(initial);
    expect(Object.isFrozen(initial)).toBe(true);
    expect(initial?.every(Object.isFrozen)).toBe(true);
    expect(shape.adjustments).not.toBe(initial);

    const noOp = packageSnapshot(pkg);
    shape.adjustments = [
      { name: 'adj1', value: 16_200_000 },
      { name: 'adj2', value: 0 },
      { name: 'adj3', value: 25_000 },
    ];
    expect(packageSnapshot(pkg)).toEqual(noOp);
    expect(slide.shapes[0]).toBe(shape);

    const invalidBefore = packageSnapshot(pkg);
    for (const invalid of [
      undefined,
      null,
      [{ name: '', value: 1 }],
      [{ name: 'adj', value: 1.5 }],
      [{ name: 'adj', value: 1 }, { name: 'adj', value: 2 }],
    ]) {
      expect(() => {
        shape.adjustments = invalid as never;
      }).toThrow();
      expect(packageSnapshot(pkg)).toEqual(invalidBefore);
    }

    slide.setShapeAdjustments(shape.id, [
      { name: 'adj1', value: 10_800_000 },
      { name: 'adj2', value: 5_400_000 },
    ]);
    expect(shape.adjustments).toEqual([
      { name: 'adj1', value: 10_800_000 },
      { name: 'adj2', value: 5_400_000 },
    ]);
    expect(shape.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent2' },
    });
    expect(shape.line?.kind).toBe('line');
    expect(shape.shadow?.kind).toBe('outer');

    expect(() => pkg.transaction(() => {
      shape.adjustments = [{ name: 'adj', value: 7 }];
      expect(shape.adjustments).toEqual([{ name: 'adj', value: 7 }]);
      throw new Error('restore shape adjustments');
    })).toThrow('restore shape adjustments');
    expect(shape.adjustments).toEqual([
      { name: 'adj1', value: 10_800_000 },
      { name: 'adj2', value: 5_400_000 },
    ]);
    expect(slide.shapes[0]).toBe(shape);

    shape.adjustments = [];
    expect(shape.adjustments).toEqual([]);
    expect(new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes))
      .toContain('<a:prstGeom prst="blockArc"><a:avLst/></a:prstGeom>');
    const cleared = packageSnapshot(pkg);
    shape.adjustments = [];
    expect(packageSnapshot(pkg)).toEqual(cleared);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedShape = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!
      .shapes.find(({ id }) => id === shape.id) as ShapeModel;
    expect(reopenedShape.adjustments).toEqual([]);
  });

  it('preserves unsupported preset shape adjustment formulas and rejects live edits', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const shape = slide.addShape('rect', { adjustments: [{ name: 'adj', value: 1 }] });
    const part = pkg.requirePart(slide.partUri);
    const complex = new TextDecoder().decode(part.bytes)
      .replace('fmla="val 1"', 'fmla="*/ 1 2 3"');
    pkg.setPart(slide.partUri, complex, part.contentType);

    expect(shape.adjustments).toBeUndefined();
    const before = pkg.requirePart(slide.partUri).bytes.slice();
    const journal = [...pkg.mutations];
    expect(() => {
      shape.adjustments = [{ name: 'adj', value: 7 }];
    }).toThrow(ModelParseError);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(before);
    expect(pkg.mutations).toEqual(journal);

    shape.setTransform({ x: inches(3) });
    expect(new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes))
      .toContain('fmla="*/ 1 2 3"');
    expect(shape.adjustments).toBeUndefined();
  });

  it('creates preset shape shadows with exact detached outer, inner, omitted, and identity state', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const omittedSlide = model.addSlide();
    const undefinedSlide = model.addSlide();
    omittedSlide.addShape('rect');
    undefinedSlide.addShape('rect', { shadow: undefined } as never);
    expect(pkg.requirePart(undefinedSlide.partUri).bytes).toEqual(
      pkg.requirePart(omittedSlide.partUri).bytes,
    );

    const slide = model.addSlide();
    const relationships = slide.relationships;
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
    const outer = slide.addShape('roundRect', { name: 'Outer shadow', shadow });
    const inner = slide.addShape('ellipse', {
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

    expect([outer.id, inner.id]).toEqual([2, 3]);
    expect(slide.shapes).toEqual([outer, inner]);
    expect(slide.shapes[0]).toBe(outer);
    expect(slide.relationships).toEqual(relationships);
    expect(readCreatedShapeShadow(model, slide, outer.id)).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '123ABC' },
      opacity: 0.42,
      blur: 7.25,
      angle: 123.4,
      distance: 5.5,
      rotateWithShape: true,
    });
    expect(readCreatedShapeShadow(model, slide, inner.id)).toEqual({
      kind: 'inner',
      color: { kind: 'scheme', value: 'accent2' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
    });
  });

  it('rejects invalid preset shape shadows without any package, identity, or ID mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const before = packageSnapshot(pkg);
    const shapes = slide.shapes;
    let calls = 0;
    const optionsGetter = Object.defineProperty({}, 'shadow', {
      enumerable: true,
      get() {
        calls += 1;
        return { kind: 'outer' };
      },
    });
    const shadowGetter = Object.defineProperty({}, 'kind', {
      enumerable: true,
      get() {
        calls += 1;
        return 'outer';
      },
    });
    for (const shadow of [
      null,
      [],
      {},
      { kind: 'none' },
      { kind: 'inner', rotateWithShape: false },
      { kind: 'outer', opacity: Number.NaN },
      { kind: 'outer', blur: 101 },
      { kind: 'outer', angle: 360 },
      { kind: 'outer', distance: 201 },
      { kind: 'outer', color: { kind: 'srgb', value: 'FFF' } },
      { kind: 'outer', type: 'outer' },
      { kind: 'outer', offset: 4 },
      { kind: 'outer', [Symbol('unsafe')]: true },
      shadowGetter,
    ]) {
      expect(() => slide.addShape('rect', { shadow } as never)).toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
      expect(slide.shapes).toEqual(shapes);
    }
    expect(() => slide.addShape('rect', optionsGetter as never)).toThrow(/data property/);
    expect(calls).toBe(0);
    expect(packageSnapshot(pkg)).toEqual(before);
    expect(slide.addShape('rect').id).toBe(2);
  });

  it('rolls back preset shape shadow bytes, journal, identity, and next ID', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const before = packageSnapshot(pkg);
    let rolledBack: ShapeModel | undefined;

    expect(() => pkg.transaction(() => {
      rolledBack = slide.addShape('rect', { shadow: { kind: 'outer' } });
      throw new Error('restore shadow creation');
    })).toThrow('restore shadow creation');

    expect(packageSnapshot(pkg)).toEqual(before);
    expect(slide.shapes).toEqual([]);
    expect(() => rolledBack!.name).toThrow(ModelParseError);
    expect(slide.addShape('rect').id).toBe(2);
  });

  it('reopens preset shape shadows in all six OOXML presentation formats', async () => {
    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const pkg = await OpcPackage.open(await modelFixture(profile.presentationContentType));
      const model = new PresentationModel(pkg);
      const slide = model.addSlide();
      const outer = slide.addShape('rect', {
        shadow: { kind: 'outer', color: { kind: 'scheme', value: 'accent4' } },
      });
      const inner = slide.addShape('ellipse', {
        shadow: { kind: 'inner', opacity: 0, blur: 0, angle: 0, distance: 0 },
      });
      const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
      const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!;
      expect(reopened.format).toBe(profile.format);
      expect(readCreatedShapeShadow(reopened, reopenedSlide, outer.id)).toEqual({
        kind: 'outer',
        color: { kind: 'scheme', value: 'accent4' },
        opacity: 0.75,
        blur: 8,
        angle: 270,
        distance: 4,
        rotateWithShape: false,
      });
      expect(readCreatedShapeShadow(reopened, reopenedSlide, inner.id)).toEqual({
        kind: 'inner',
        color: { kind: 'srgb', value: '000000' },
        opacity: 0,
        blur: 0,
        angle: 0,
        distance: 0,
      });
    }
  });

  it('supports the shape shadow lifecycle with exact no-ops, edits, clear, and rollback', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const shape = slide.addShape('rect', {
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

    const initial = shape.shadow;
    expect(initial).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '123ABC' },
      opacity: 0.42,
      blur: 7.25,
      angle: 123.4,
      distance: 5.5,
      rotateWithShape: true,
    });
    expect(Object.isFrozen(initial)).toBe(true);
    expect(Object.isFrozen(initial?.color)).toBe(true);
    expect(slide.getShapeShadow(shape.id)).toEqual(initial);

    const noOp = packageSnapshot(pkg);
    shape.shadow = {
      kind: 'outer',
      color: { kind: 'srgb', value: '123ABC' },
      opacity: 0.42,
      blur: 7.25,
      angle: 123.4,
      distance: 5.5,
      rotateWithShape: true,
    };
    expect(packageSnapshot(pkg)).toEqual(noOp);
    expect(slide.shapes[0]).toBe(shape);

    const invalidBefore = packageSnapshot(pkg);
    for (const invalid of [
      { kind: 'none' },
      { kind: 'inner', rotateWithShape: false },
      { kind: 'outer', opacity: Number.NaN },
      { kind: 'outer', angle: 360 },
      { kind: 'outer', color: { kind: 'srgb', value: 'FFF' } },
      { kind: 'outer', offset: 4 },
    ]) {
      expect(() => {
        shape.shadow = invalid as never;
      }).toThrow();
      expect(packageSnapshot(pkg)).toEqual(invalidBefore);
      expect(slide.shapes[0]).toBe(shape);
    }

    slide.setShapeShadow(shape.id, {
      kind: 'inner',
      color: { kind: 'scheme', value: 'accent3' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
    });
    expect(shape.shadow).toEqual({
      kind: 'inner',
      color: { kind: 'scheme', value: 'accent3' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
    });

    expect(() => pkg.transaction(() => {
      shape.shadow = { kind: 'outer', color: { kind: 'scheme', value: 'accent5' } };
      expect(shape.shadow?.kind).toBe('outer');
      throw new Error('restore shape shadow');
    })).toThrow('restore shape shadow');
    expect(shape.shadow?.kind).toBe('inner');
    expect(slide.shapes[0]).toBe(shape);

    shape.shadow = undefined;
    expect(shape.shadow).toBeUndefined();
    const cleared = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(cleared).toContain('<a:effectLst></a:effectLst>');
    expect(cleared).not.toContain('innerShdw');
    const clearedNoOp = packageSnapshot(pkg);
    shape.shadow = undefined;
    expect(packageSnapshot(pkg)).toEqual(clearedNoOp);
  });

  it('rejects malformed shape shadow lifecycle state without package or identity mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const shape = slide.addShape('rect', { shadow: { kind: 'outer' } });
    const part = pkg.requirePart(slide.partUri);
    const malformed = new TextDecoder().decode(part.bytes).replace(
      '<a:effectLst><a:outerShdw',
      '<a:effectLst><a:reflection/><a:outerShdw',
    );
    pkg.setPart(slide.partUri, malformed, part.contentType);
    const before = packageSnapshot(pkg);

    expect(shape.shadow).toBeUndefined();
    expect(() => {
      shape.shadow = { kind: 'inner' };
    }).toThrow(ModelParseError);
    expect(packageSnapshot(pkg)).toEqual(before);
    expect(slide.shapes[0]).toBe(shape);
  });

  it('creates preset shape hyperlinks with exact URL, slide, self, and tooltip semantics', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const first = model.addSlide();
    const second = model.addSlide();
    pkg.addRelationship(first.partUri, {
      id: 'rId2',
      type: 'urn:example:relationships/opaque',
      target: 'https://opaque.example/two',
      targetMode: 'External',
    });
    pkg.addRelationship(first.partUri, {
      id: 'rId4',
      type: 'urn:example:relationships/opaque',
      target: 'https://opaque.example/four',
      targetMode: 'External',
    });
    const websiteInput = {
      url: 'https://example.com?a=1&b=2',
      tooltip: 'Visit & learn',
    };
    const nextInput: { slide: number; tooltip?: string } = {
      slide: model.slides.length,
      tooltip: '',
    };

    const website = first.addShape('rect', {
      name: 'Website',
      hyperlink: websiteInput,
    });
    const next = first.addShape('actionButtonForwardNext', {
      name: 'Next slide',
      hyperlink: nextInput,
    });
    const self = first.addShape('actionButtonHome', {
      hyperlink: { slide: model.slides.indexOf(first) + 1 },
    });
    websiteInput.url = 'https://changed.example';
    websiteInput.tooltip = 'Changed';
    nextInput.slide = 1;
    delete nextInput.tooltip;

    expect([website.id, next.id, self.id]).toEqual([2, 3, 4]);
    expect(first.shapes).toEqual([website, next, self]);
    expect(first.shapes[0]).toBe(website);
    expect(readCreatedShapeHyperlink(model, first, website.id)).toEqual({
      url: 'https://example.com?a=1&b=2',
      tooltip: 'Visit & learn',
    });
    expect(readCreatedShapeHyperlink(model, first, next.id)).toEqual({
      slide: model.slides.indexOf(second) + 1,
      tooltip: '',
    });
    const selfLink = readCreatedShapeHyperlink(model, first, self.id);
    expect(selfLink).toEqual({ slide: model.slides.indexOf(first) + 1 });
    expect(Object.hasOwn(selfLink!, 'tooltip')).toBe(false);

    const relationships = first.relationships;
    expect(relationships.find(({ type }) => type === HYPERLINK_RELATIONSHIP)).toMatchObject({
      target: 'https://example.com?a=1&b=2',
      targetMode: 'External',
    });
    expect(relationships.filter(({ type }) => type === SLIDE_RELATIONSHIP)).toEqual([
      expect.objectContaining({
        targetMode: 'Internal',
        resolvedTarget: second.partUri,
      }),
      expect.objectContaining({
        targetMode: 'Internal',
        resolvedTarget: first.partUri,
      }),
    ]);
    const source = new TextDecoder().decode(pkg.requirePart(first.partUri).bytes);
    const relationshipSource = new TextDecoder().decode(
      pkg.requirePart(relationshipPartUri(first.partUri)).bytes,
    );
    expect(source).toContain('r:id="rId1" tooltip="Visit &amp; learn"');
    expect(source).toContain('r:id="rId3" tooltip="" action="ppaction://hlinksldjump"');
    expect(source).toContain('r:id="rId5" action="ppaction://hlinksldjump"');
    expect(source).toContain('tooltip="Visit &amp; learn"');
    expect(source).toContain('tooltip="" action="ppaction://hlinksldjump"');
    expect(source.match(/ppaction:\/\/hlinksldjump/g)).toHaveLength(2);
    expect(relationshipSource).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ' +
      'Target="https://example.com?a=1&amp;b=2" TargetMode="External"',
    );
  });

  it('preserves omitted bytes and rejects invalid preset shape hyperlinks without mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const omittedSlide = model.addSlide();
    const undefinedSlide = model.addSlide();
    omittedSlide.addShape('rect');
    undefinedSlide.addShape('rect', { hyperlink: undefined } as never);
    expect(pkg.requirePart(undefinedSlide.partUri).bytes).toEqual(
      pkg.requirePart(omittedSlide.partUri).bytes,
    );
    expect(undefinedSlide.relationships).toEqual(omittedSlide.relationships);

    const slide = model.addSlide();
    const before = packageSnapshot(pkg);
    const shapes = [...slide.shapes];
    let calls = 0;
    const hyperlinkAccessors = (['url', 'slide', 'tooltip'] as const).map((key) =>
      Object.defineProperty({}, key, {
        enumerable: true,
        get() {
          calls += 1;
          return key === 'slide' ? 1 : 'https://example.com';
        },
      }));
    const optionsGetter = Object.defineProperty({}, 'hyperlink', {
      enumerable: true,
      get() {
        calls += 1;
        return { url: 'https://example.com' };
      },
    });
    for (const hyperlink of [
      null,
      false,
      [],
      {},
      { url: 'https://example.com', slide: 1 },
      { url: '' },
      { url: 42 },
      { slide: 0 },
      { slide: -1 },
      { slide: 1.5 },
      { slide: Number.MAX_SAFE_INTEGER + 1 },
      { slide: model.slides.length + 1 },
      { url: 'bad\u0000url' },
      { url: 'https://example.com', tooltip: 'bad\u0000tooltip' },
      { url: 'https://example.com', _rId: 'rId9' },
      { url: 'https://example.com', [Symbol('unsafe')]: true },
      ...hyperlinkAccessors,
    ]) {
      expect(() => slide.addShape('rect', { hyperlink } as never)).toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
      expect(slide.shapes).toEqual(shapes);
    }
    expect(() => slide.addShape('rect', optionsGetter as never)).toThrow(/data property/);
    expect(calls).toBe(0);
    expect(packageSnapshot(pkg)).toEqual(before);
    expect(slide.addShape('rect').id).toBe(2);
  });

  it('rolls back preset shape hyperlink XML, relationship graph, journal, identity, and next ID', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const before = packageSnapshot(pkg);
    let rolledBack: ShapeModel | undefined;

    expect(() => pkg.transaction(() => {
      rolledBack = slide.addShape('rect', {
        hyperlink: { url: 'https://example.com' },
      });
      throw new Error('restore hyperlink creation');
    })).toThrow('restore hyperlink creation');

    expect(packageSnapshot(pkg)).toEqual(before);
    expect(slide.shapes).toEqual([]);
    expect(() => rolledBack!.name).toThrow(ModelParseError);
    expect(slide.addShape('rect').id).toBe(2);
  });

  it('supports the shape hyperlink lifecycle with exact no-ops and relationship reuse', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const source = model.addSlide();
    const target = model.addSlide();
    const other = model.addSlide();
    const shape = source.addShape('rect', {
      hyperlink: { url: 'https://example.com', tooltip: 'Visit' },
    });

    const initial = shape.hyperlink;
    expect(initial).toEqual({ url: 'https://example.com', tooltip: 'Visit' });
    expect(Object.isFrozen(initial)).toBe(true);
    expect(source.getShapeHyperlink(shape.id)).toEqual(initial);
    const noOp = packageSnapshot(pkg);
    shape.hyperlink = { url: 'https://example.com', tooltip: 'Visit' };
    expect(packageSnapshot(pkg)).toEqual(noOp);
    expect(source.shapes[0]).toBe(shape);

    const relationshipUri = relationshipPartUri(source.partUri);
    const relationshipBytes = pkg.requirePart(relationshipUri).bytes.slice();
    const initialRelationship = source.relationships.find(
      ({ type }) => type === HYPERLINK_RELATIONSHIP,
    )!;
    shape.hyperlink = { url: 'https://example.com', tooltip: '' };
    expect(shape.hyperlink).toEqual({ url: 'https://example.com', tooltip: '' });
    expect(pkg.requirePart(relationshipUri).bytes).toEqual(relationshipBytes);

    shape.hyperlink = { url: 'mailto:test@example.com' };
    expect(source.relationships.find(({ id }) => id === initialRelationship.id)).toMatchObject({
      type: HYPERLINK_RELATIONSHIP,
      target: 'mailto:test@example.com',
      targetMode: 'External',
    });
    shape.hyperlink = { slide: model.slides.indexOf(target) + 1, tooltip: '' };
    expect(shape.hyperlink).toEqual({
      slide: model.slides.indexOf(target) + 1,
      tooltip: '',
    });
    expect(source.relationships.find(({ id }) => id === initialRelationship.id)).toMatchObject({
      type: SLIDE_RELATIONSHIP,
      targetMode: 'Internal',
      resolvedTarget: target.partUri,
    });

    model.moveSlide(model.slides.indexOf(target), 0);
    expect(shape.hyperlink).toEqual({ slide: 1, tooltip: '' });
    expect(source.relationships.find(({ id }) => id === initialRelationship.id)?.resolvedTarget)
      .toBe(target.partUri);
    shape.hyperlink = { slide: model.slides.indexOf(other) + 1 };
    expect(shape.hyperlink).toEqual({ slide: model.slides.indexOf(other) + 1 });
    shape.hyperlink = { slide: model.slides.indexOf(source) + 1 };
    expect(shape.hyperlink).toEqual({ slide: model.slides.indexOf(source) + 1 });

    const beforeRollback = packageSnapshot(pkg);
    expect(() => pkg.transaction(() => {
      shape.hyperlink = { url: 'https://rollback.example', tooltip: 'Rollback' };
      shape.hyperlink = undefined;
      throw new Error('restore shape hyperlink lifecycle');
    })).toThrow('restore shape hyperlink lifecycle');
    expect(packageSnapshot(pkg)).toEqual(beforeRollback);
    expect(shape.hyperlink).toEqual({ slide: model.slides.indexOf(source) + 1 });

    shape.hyperlink = undefined;
    expect(shape.hyperlink).toBeUndefined();
    expect(source.getShapeHyperlink(shape.id)).toBeUndefined();
    expect(source.relationships.some(({ id }) => id === initialRelationship.id)).toBe(false);
    expect(source.shapes[0]).toBe(shape);
  });

  it('clone-on-writes and garbage-collects shared shape hyperlink relationships', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const source = model.addSlide();
    const first = source.addShape('rect', {
      hyperlink: { url: 'https://shared.example' },
    });
    const second = source.addShape('ellipse', {
      hyperlink: { url: 'https://temporary.example' },
    });
    const [sharedRelationship, temporaryRelationship] = source.relationships.filter(
      ({ type }) => type === HYPERLINK_RELATIONSHIP,
    );
    expect(sharedRelationship).toBeDefined();
    expect(temporaryRelationship).toBeDefined();
    const part = pkg.requirePart(source.partUri);
    const sharedTextClick = `<a:hlinkClick r:id="${sharedRelationship!.id}"/>`;
    const sharedSource = new TextDecoder().decode(part.bytes)
      .replace(`r:id="${temporaryRelationship!.id}"`, `r:id="${sharedRelationship!.id}"`)
      .replace(
        '</p:spTree>',
        `<x:opaque xmlns:x="urn:test"><a:rPr>${sharedTextClick}</a:rPr></x:opaque>` +
        '</p:spTree>',
      );
    pkg.setPart(source.partUri, sharedSource, part.contentType);
    pkg.removeRelationship(source.partUri, temporaryRelationship!.id);

    expect(first.hyperlink).toEqual({ url: 'https://shared.example' });
    expect(second.hyperlink).toEqual({ url: 'https://shared.example' });
    first.hyperlink = { url: 'https://first.example', tooltip: 'First' };
    expect(first.hyperlink).toEqual({ url: 'https://first.example', tooltip: 'First' });
    expect(second.hyperlink).toEqual({ url: 'https://shared.example' });
    expect(source.relationships.filter(({ type }) => type === HYPERLINK_RELATIONSHIP))
      .toHaveLength(2);

    const oldRelationshipBytes = pkg.requirePart(relationshipPartUri(source.partUri)).bytes.slice();
    first.hyperlink = { url: 'https://first.example', tooltip: '' };
    expect(pkg.requirePart(relationshipPartUri(source.partUri)).bytes)
      .toEqual(oldRelationshipBytes);
    first.hyperlink = undefined;
    expect(source.relationships.filter(({ type }) => type === HYPERLINK_RELATIONSHIP))
      .toHaveLength(1);
    second.hyperlink = undefined;
    expect(source.relationships.find(({ id }) => id === sharedRelationship!.id)).toBeDefined();

    const withoutTextReference = new TextDecoder().decode(pkg.requirePart(source.partUri).bytes)
      .replace(sharedTextClick, '');
    pkg.setPart(source.partUri, withoutTextReference, part.contentType);
    second.hyperlink = { url: 'https://unique.example' };
    const uniqueRelationship = source.relationships.find(
      ({ type, target }) => type === HYPERLINK_RELATIONSHIP && target === 'https://unique.example',
    )!;
    second.hyperlink = undefined;
    expect(source.relationships.some(({ id }) => id === uniqueRelationship.id)).toBe(false);
  });

  it('rejects unsupported shape hyperlink state and cleans target-slide DrawingML references', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const source = model.addSlide();
    const target = model.addSlide();
    const linked = source.addShape('rect', {
      hyperlink: { slide: model.slides.indexOf(target) + 1 },
    });
    const external = source.addShape('ellipse', {
      hyperlink: { url: 'https://keep.example' },
    });
    const targetRelationship = source.relationships.find(
      ({ type, resolvedTarget }) => type === SLIDE_RELATIONSHIP && resolvedTarget === target.partUri,
    )!;
    const part = pkg.requirePart(source.partUri);
    const decorated = new TextDecoder().decode(part.bytes)
      .replace(
        '</p:cNvPr>',
        `<a:hlinkHover r:id="${targetRelationship.id}"/></p:cNvPr>`,
      )
      .replace(
        '</p:spTree>',
        `<x:opaque xmlns:x="urn:test"><a:rPr><a:hlinkClick r:id="${
          targetRelationship.id
        }"/></a:rPr></x:opaque></p:spTree>`,
      );
    pkg.setPart(source.partUri, decorated, part.contentType);

    const beforeRollback = packageSnapshot(pkg);
    expect(() => pkg.transaction(() => {
      model.deleteSlide(model.slides.indexOf(target));
      throw new Error('restore hyperlink target');
    })).toThrow('restore hyperlink target');
    expect(packageSnapshot(pkg)).toEqual(beforeRollback);
    expect(linked.hyperlink).toEqual({ slide: model.slides.indexOf(target) + 1 });

    model.deleteSlide(model.slides.indexOf(target));
    const cleaned = new TextDecoder().decode(pkg.requirePart(source.partUri).bytes);
    expect(cleaned).not.toContain(targetRelationship.id);
    expect(cleaned).not.toContain('<a:hlinkHover');
    expect(linked.hyperlink).toBeUndefined();
    expect(external.hyperlink).toEqual({ url: 'https://keep.example' });
    expect(source.relationships.some(({ id }) => id === targetRelationship.id)).toBe(false);

    const dangling = source.addShape('diamond', {
      hyperlink: { url: 'https://dangling.example' },
    });
    const danglingRelationship = source.relationships.find(
      ({ type, target: relationshipTarget }) =>
        type === HYPERLINK_RELATIONSHIP && relationshipTarget === 'https://dangling.example',
    )!;
    pkg.removeRelationship(source.partUri, danglingRelationship.id);
    expect(dangling.hyperlink).toBeUndefined();
    const unsupported = packageSnapshot(pkg);
    expect(() => {
      dangling.hyperlink = { url: 'https://replacement.example' };
    }).toThrow(ModelParseError);
    expect(() => {
      dangling.hyperlink = undefined;
    }).toThrow(ModelParseError);
    expect(packageSnapshot(pkg)).toEqual(unsupported);

    expect(() => {
      external.hyperlink = { slide: model.slides.length + 1 };
    }).toThrow(RangeError);
    expect(packageSnapshot(pkg)).toEqual(unsupported);

    const currentPart = pkg.requirePart(source.partUri);
    pkg.setPart(source.partUri, '<broken', currentPart.contentType);
    const malformed = packageSnapshot(pkg);
    expect(() => {
      source.setShapeHyperlink(dangling.id, { slide: 0 } as never);
    }).toThrow(TypeError);
    expect(packageSnapshot(pkg)).toEqual(malformed);
  });

  it('creates preset shape fills with detached strict values through the live model', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const runtimeUndefined: ShapeFill | undefined = undefined;
    const mutableFill: {
      kind: 'solid';
      color: { kind: 'srgb'; value: string };
      transparency: number;
    } = {
      kind: 'solid',
      color: { kind: 'srgb', value: '#ff0000' },
      transparency: 50,
    };

    const omitted = slide.addShape('rect');
    const undefinedFill = slide.addShape('ellipse', { fill: runtimeUndefined } as never);
    const none = slide.addShape('star5', { fill: { kind: 'none' } });
    const red = slide.addShape('diamond', { fill: mutableFill });
    const themed = slide.addShape('hexagon', {
      fill: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 25,
      },
    });
    mutableFill.color.value = '000000';
    mutableFill.transparency = 1;

    expect([omitted, undefinedFill, none, red, themed].every(
      (shape) => shape instanceof ShapeModel,
    )).toBe(true);
    expect(slide.shapes).toEqual([omitted, undefinedFill, none, red, themed]);
    expect(slide.shapes[3]).toBe(red);

    const xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml).toContain(
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/>',
    );
    expect(xml).toContain(
      '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:noFill/><a:ln/>',
    );
    expect(xml).toContain(
      '<a:prstGeom prst="star5"><a:avLst/></a:prstGeom><a:noFill/><a:ln/>',
    );
    expect(xml).toContain(
      '<a:prstGeom prst="diamond"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/>' +
      '</a:srgbClr></a:solidFill><a:ln/>',
    );
    expect(xml).toContain(
      '<a:prstGeom prst="hexagon"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:schemeClr val="accent2"><a:alpha val="75000"/>' +
      '</a:schemeClr></a:solidFill><a:ln/>',
    );
  });

  it('rejects invalid preset shape fills before mutating package state', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const partBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const parts = pkg.parts.map(({ uri, contentType, bytes }) => ({
      uri,
      contentType,
      bytes: bytes.slice(),
    }));
    const relationships = slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }));
    const shapes = [...slide.shapes];
    const journal = [...pkg.mutations];
    let accessorCalls = 0;
    const inherited = Object.create({ kind: 'none' });
    const fillGetter = Object.defineProperty({}, 'kind', {
      enumerable: true,
      get() {
        accessorCalls += 1;
        return 'none';
      },
    });
    const optionsGetter = Object.defineProperty({}, 'fill', {
      enumerable: true,
      get() {
        accessorCalls += 1;
        return { kind: 'none' };
      },
    });
    const invalidFills = [
      null,
      false,
      [],
      new Date(),
      inherited,
      fillGetter,
      { kind: 'gradient' },
      { kind: 'none', color: undefined },
      { kind: 'solid' },
      { kind: 'solid', color: { kind: 'srgb', value: 'FFF' } },
      { kind: 'solid', color: { kind: 'scheme', value: 'unknown' } },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFFFFF' },
        transparency: Number.NaN,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFFFFF' },
        transparency: -1,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFFFFF' },
        transparency: 101,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFFFFF' },
        alpha: 40,
      },
      { kind: 'none', type: 'none' },
      { kind: 'none', [Symbol('unsafe')]: true },
    ];
    for (const fill of invalidFills) {
      expect(() => slide.addShape('rect', { fill } as never)).toThrow();
    }
    expect(() => slide.addShape('rect', optionsGetter as never)).toThrow(/data property/);

    expect(accessorCalls).toBe(0);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(partBytes);
    expect(pkg.parts.map(({ uri, contentType, bytes }) => ({
      uri,
      contentType,
      bytes,
    }))).toEqual(parts);
    expect(slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }))).toEqual(relationships);
    expect(slide.shapes).toEqual(shapes);
    expect(pkg.mutations).toEqual(journal);
  });

  it('creates preset shape lines with detached strict values through the live model', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const runtimeUndefined: ShapeLine | undefined = undefined;
    const dash: ShapeLineDash = 'dashDot';
    const mutableLine: {
      kind: 'line';
      color: { kind: 'srgb'; value: string };
      transparency: number;
      width: number;
      dash: ShapeLineDash;
    } = {
      kind: 'line',
      color: { kind: 'srgb', value: '#ff0000' },
      transparency: 50,
      width: 2.5,
      dash,
    };

    const omitted = slide.addShape('rect');
    const undefinedLine = slide.addShape('ellipse', { line: runtimeUndefined } as never);
    const none = slide.addShape('star5', { line: { kind: 'none' } });
    const red = slide.addShape('diamond', { line: mutableLine });
    const themed = slide.addShape('hexagon', {
      line: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 25,
        width: 0,
        dash: 'sysDot',
      },
    });
    mutableLine.color.value = '000000';
    mutableLine.transparency = 1;
    mutableLine.width = 9;
    mutableLine.dash = 'solid';

    expect([omitted, undefinedLine, none, red, themed].every(
      (shape) => shape instanceof ShapeModel,
    )).toBe(true);
    expect(slide.shapes).toEqual([omitted, undefinedLine, none, red, themed]);
    expect(slide.shapes[3]).toBe(red);

    const xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect((xml.match(/<a:ln\/>/g) ?? [])).toHaveLength(2);
    expect(xml).toContain(
      '<a:prstGeom prst="star5"><a:avLst/></a:prstGeom>' +
      '<a:noFill/><a:ln><a:noFill/></a:ln>',
    );
    expect(xml).toContain(
      '<a:prstGeom prst="diamond"><a:avLst/></a:prstGeom><a:noFill/>' +
      '<a:ln w="31750"><a:solidFill><a:srgbClr val="FF0000">' +
      '<a:alpha val="50000"/></a:srgbClr></a:solidFill>' +
      '<a:prstDash val="dashDot"/></a:ln>',
    );
    expect(xml).toContain(
      '<a:prstGeom prst="hexagon"><a:avLst/></a:prstGeom><a:noFill/>' +
      '<a:ln w="0"><a:solidFill><a:schemeClr val="accent2">' +
      '<a:alpha val="75000"/></a:schemeClr></a:solidFill>' +
      '<a:prstDash val="sysDot"/></a:ln>',
    );

    // @ts-expect-error preset shape line dash excludes arbitrary DrawingML tokens
    const invalidDash: ShapeLineDash = 'dot';
    // @ts-expect-error a solid line must use the native kind discriminator
    const invalidLine: ShapeLine = { kind: 'solid', color: { kind: 'srgb', value: 'FFFFFF' } };
    void [invalidDash, invalidLine];
  });

  it('rejects invalid preset shape lines before mutating package state', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const partBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const parts = pkg.parts.map(({ uri, contentType, bytes }) => ({
      uri,
      contentType,
      bytes: bytes.slice(),
    }));
    const relationships = slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }));
    const shapes = [...slide.shapes];
    const journal = [...pkg.mutations];
    let accessorCalls = 0;
    const inherited = Object.create({ kind: 'none' });
    const lineGetter = Object.defineProperty({}, 'kind', {
      enumerable: true,
      get() {
        accessorCalls += 1;
        return 'none';
      },
    });
    const optionsGetter = Object.defineProperty({}, 'line', {
      enumerable: true,
      get() {
        accessorCalls += 1;
        return { kind: 'none' };
      },
    });
    const invalidLines = [
      null,
      false,
      [],
      new Date(),
      inherited,
      lineGetter,
      { kind: 'solid' },
      { kind: 'none', width: undefined },
      { kind: 'line' },
      { kind: 'line', color: { kind: 'srgb', value: 'FFF' } },
      { kind: 'line', color: { kind: 'scheme', value: 'unknown' } },
      { kind: 'line', color: { kind: 'srgb', value: 'FFFFFF' }, transparency: -1 },
      { kind: 'line', color: { kind: 'srgb', value: 'FFFFFF' }, width: 1_585 },
      { kind: 'line', color: { kind: 'srgb', value: 'FFFFFF' }, dash: 'dot' },
      { type: 'none' },
      { color: 'FFFFFF', dashType: 'dash' },
      { kind: 'line', color: { kind: 'srgb', value: 'FFFFFF' }, alpha: 40 },
      { kind: 'line', color: { kind: 'srgb', value: 'FFFFFF' }, lineDash: 'dash' },
      { kind: 'none', [Symbol('unsafe')]: true },
    ];
    for (const line of invalidLines) {
      expect(() => slide.addShape('rect', { line } as never)).toThrow();
    }
    expect(() => slide.addShape('rect', optionsGetter as never)).toThrow(/data property/);

    expect(accessorCalls).toBe(0);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(partBytes);
    expect(pkg.parts.map(({ uri, contentType, bytes }) => ({
      uri,
      contentType,
      bytes,
    }))).toEqual(parts);
    expect(slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }))).toEqual(relationships);
    expect(slide.shapes).toEqual(shapes);
    expect(pkg.mutations).toEqual(journal);
  });

  it('creates preset shape arrows with detached strict values through the live model', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const runtimeUndefined: ShapeArrows | undefined = undefined;
    const mutableArrows: { begin: ShapeArrowType; end: ShapeArrowType } = {
      begin: 'triangle',
      end: 'arrow',
    };

    const omitted = slide.addShape('line');
    const undefinedArrows = slide.addShape('lineInv', {
      arrows: runtimeUndefined,
    } as never);
    const empty = slide.addShape('line', { arrows: {} });
    const both = slide.addShape('line', { arrows: mutableArrows });
    const arrowOnly = slide.addShape('lineInv', { arrows: { end: 'stealth' } });
    const noneLine = slide.addShape('line', {
      line: { kind: 'none' },
      arrows: { begin: 'none', end: 'oval' },
    });
    const styled = slide.addShape('line', {
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
        width: 2.5,
        dash: 'dashDot',
      },
      arrows: { begin: 'diamond', end: 'triangle' },
    });
    mutableArrows.begin = 'none';
    mutableArrows.end = 'none';

    expect([omitted, undefinedArrows, empty, both, arrowOnly, noneLine, styled].every(
      (shape) => shape instanceof ShapeModel,
    )).toBe(true);
    expect(slide.shapes).toEqual([
      omitted,
      undefinedArrows,
      empty,
      both,
      arrowOnly,
      noneLine,
      styled,
    ]);
    expect(slide.shapes[3]).toBe(both);

    const xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect((xml.match(/<a:ln\/>/g) ?? [])).toHaveLength(3);
    expect(xml).toContain(
      '<a:ln><a:headEnd type="triangle"/><a:tailEnd type="arrow"/></a:ln>',
    );
    expect(xml).toContain('<a:ln><a:tailEnd type="stealth"/></a:ln>');
    expect(xml).toContain(
      '<a:ln><a:noFill/><a:headEnd type="none"/><a:tailEnd type="oval"/></a:ln>',
    );
    expect(xml).toContain(
      '<a:ln w="31750"><a:solidFill><a:srgbClr val="112233"/>' +
      '</a:solidFill><a:prstDash val="dashDot"/>' +
      '<a:headEnd type="diamond"/><a:tailEnd type="triangle"/></a:ln>',
    );
    expect(xml).not.toContain('333333');
  });

  it('rejects invalid preset shape arrows before mutating package state', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const partBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const parts = pkg.parts.map(({ uri, contentType, bytes }) => ({
      uri,
      contentType,
      bytes: bytes.slice(),
    }));
    const relationships = slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }));
    const shapes = [...slide.shapes];
    const journal = [...pkg.mutations];
    let accessorCalls = 0;
    const arrowsGetter = Object.defineProperty({}, 'begin', {
      enumerable: true,
      get() {
        accessorCalls += 1;
        return 'arrow';
      },
    });
    const optionsGetter = Object.defineProperty({}, 'arrows', {
      enumerable: true,
      get() {
        accessorCalls += 1;
        return { begin: 'arrow' };
      },
    });
    for (const arrows of [
      null,
      false,
      [],
      new Date(),
      Object.create({ begin: 'arrow' }),
      arrowsGetter,
      { begin: '' },
      { begin: 'Arrow' },
      { begin: 'bogus' },
      { end: null },
      { beginArrowType: 'arrow' },
      { endArrowType: 'arrow' },
      { lineHead: 'arrow' },
      { lineTail: 'arrow' },
      { begin: 'arrow', [Symbol('unsafe')]: true },
    ]) {
      expect(() => slide.addShape('line', { arrows } as never)).toThrow();
    }
    expect(() => slide.addShape('line', optionsGetter as never)).toThrow(/data property/);
    expect(() => slide.addShape('line', { beginArrowType: 'arrow' } as never))
      .toThrow(/unsupported property/);

    expect(accessorCalls).toBe(0);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(partBytes);
    expect(pkg.parts.map(({ uri, contentType, bytes }) => ({ uri, contentType, bytes })))
      .toEqual(parts);
    expect(slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }))).toEqual(relationships);
    expect(slide.shapes).toEqual(shapes);
    expect(pkg.mutations).toEqual(journal);
  });

  it('reads and edits direct shape arrows through stable live models', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const shape = slide.addShape('line', {
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
        width: 2.5,
        dash: 'dashDot',
      },
      arrows: { begin: 'triangle', end: 'arrow' },
    });
    const text = slide.addText('Editable text arrows');

    expect(shape.arrows).toEqual({ begin: 'triangle', end: 'arrow' });
    expect(text.arrows).toBeUndefined();
    const first = shape.arrows;
    const second = shape.arrows;
    expect(first).not.toBe(second);
    expect(Object.isFrozen(first)).toBe(true);
    expect(() => {
      if (first) (first as { begin?: ShapeArrowType }).begin = 'none';
    }).toThrow(TypeError);
    expect(shape.arrows).toEqual({ begin: 'triangle', end: 'arrow' });

    const noOpBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...pkg.mutations];
    shape.arrows = { begin: 'triangle', end: 'arrow' };
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(noOpBytes);
    expect(pkg.mutations).toEqual(noOpJournal);

    const part = pkg.requirePart(slide.partUri);
    const sourceXml = new TextDecoder().decode(part.bytes);
    const advancedXml = sourceXml.replace(
      '<a:solidFill><a:srgbClr val="112233"/></a:solidFill>' +
      '<a:prstDash val="dashDot"/><a:headEnd type="triangle"/>' +
      '<a:tailEnd type="arrow"/></a:ln>',
      '<a:gradFill><a:gsLst/></a:gradFill>' +
      '<a:custDash><a:ds d="1" sp="1"/></a:custDash><a:round/>' +
      '<a:headEnd type="triangle" w="lg" len="sm"/>' +
      '<a:tailEnd type="arrow" w="med" len="med"/>' +
      '<a:extLst><a:ext uri="urn:arrows"><x:keep xmlns:x="urn:arrows"/>' +
      '</a:ext></a:extLst></a:ln>',
    );
    expect(advancedXml).not.toBe(sourceXml);
    pkg.setPart(slide.partUri, advancedXml, part.contentType);
    expect(shape.line).toBeUndefined();
    expect(shape.arrows).toEqual({ begin: 'triangle', end: 'arrow' });

    shape.arrows = { begin: 'diamond', end: 'oval' };
    let xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml).toContain('<a:gradFill><a:gsLst/></a:gradFill>');
    expect(xml).toContain('<a:custDash><a:ds d="1" sp="1"/></a:custDash><a:round/>');
    expect(xml).toContain('<a:headEnd type="diamond" w="lg" len="sm"/>');
    expect(xml).toContain('<a:tailEnd type="oval" w="med" len="med"/>');
    expect(xml).toContain('<x:keep xmlns:x="urn:arrows"/>');

    shape.arrows = { begin: 'stealth' };
    expect(shape.arrows).toEqual({ begin: 'stealth' });
    xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml).toContain('<a:headEnd type="stealth" w="lg" len="sm"/>');
    expect(xml).not.toContain('<a:tailEnd');

    shape.arrows = { begin: 'none', end: 'triangle' };
    shape.line = undefined;
    expect(shape.line).toBeUndefined();
    expect(shape.arrows).toEqual({ begin: 'none', end: 'triangle' });
    xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml).toContain(
      '<a:ln><a:round/><a:headEnd type="none" w="lg" len="sm"/>' +
      '<a:tailEnd type="triangle"/>',
    );
    expect(xml).toContain('<x:keep xmlns:x="urn:arrows"/>');

    shape.arrows = undefined;
    expect(shape.arrows).toBeUndefined();
    xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml).not.toContain('<a:headEnd');
    expect(xml).not.toContain('<a:tailEnd');
    expect(xml).toContain('<a:round/><a:extLst>');

    text.arrows = { begin: 'diamond', end: 'oval' };
    expect(text.arrows).toEqual({ begin: 'diamond', end: 'oval' });
    expect(text.line).toEqual({ kind: 'none' });
    expect(slide.shapes[0]).toBe(shape);
    expect(slide.shapes[1]).toBe(text);

    const invalidBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const invalidJournal = [...pkg.mutations];
    let getterInvoked = false;
    const accessor = Object.defineProperty({}, 'begin', {
      enumerable: true,
      get() {
        getterInvoked = true;
        return 'arrow';
      },
    });
    expect(() => {
      shape.arrows = accessor as never;
    }).toThrow(TypeError);
    expect(() => {
      shape.arrows = { begin: 'open' } as never;
    }).toThrow(TypeError);
    expect(getterInvoked).toBe(false);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(invalidBytes);
    expect(pkg.mutations).toEqual(invalidJournal);
  });

  it('rejects malformed existing arrows through the live editor with zero mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const shape = slide.addShape('line', { arrows: { begin: 'triangle' } });
    const part = pkg.requirePart(slide.partUri);
    const malformed = new TextDecoder().decode(part.bytes).replace(
      '<a:headEnd type="triangle"/>',
      '<a:headEnd type="triangle" w="xl"/>',
    );
    pkg.setPart(slide.partUri, malformed, part.contentType);
    expect(shape.arrows).toBeUndefined();

    const bytes = pkg.requirePart(slide.partUri).bytes.slice();
    const parts = pkg.parts.map(({ uri, contentType, bytes: value }) => ({
      uri,
      contentType,
      bytes: value.slice(),
    }));
    const relationships = slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }));
    const journal = [...pkg.mutations];
    expect(() => {
      shape.arrows = { begin: 'arrow' };
    }).toThrow(ModelParseError);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(bytes);
    expect(pkg.parts.map(({ uri, contentType, bytes: value }) => ({
      uri,
      contentType,
      bytes: value,
    }))).toEqual(parts);
    expect(slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }))).toEqual(relationships);
    expect(pkg.mutations).toEqual(journal);
    expect(slide.shapes[0]).toBe(shape);
  });

  it('reads and edits direct shape lines through stable live models', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const shape = slide.addShape('rect', {
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        transparency: 25,
        width: 2.5,
        dash: 'dashDot',
      },
    });
    const text = slide.addText('Keep text line');

    expect(shape.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: 'FF0000' },
      transparency: 25,
      width: 2.5,
      dash: 'dashDot',
    });
    expect(text.line).toEqual({ kind: 'none' });
    const first = shape.line;
    const second = shape.line;
    expect(first).not.toBe(second);
    if (first?.kind === 'line' && second?.kind === 'line') {
      expect(first.color).not.toBe(second.color);
      (first.color as { value: string }).value = '000000';
    }
    expect(shape.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: 'FF0000' },
      transparency: 25,
      width: 2.5,
      dash: 'dashDot',
    });

    const beforeNoOp = pkg.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...pkg.mutations];
    shape.line = {
      kind: 'line',
      color: { kind: 'srgb', value: 'FF0000' },
      transparency: 25,
      width: 2.5,
      dash: 'dashDot',
    };
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeNoOp);
    expect(pkg.mutations).toEqual(noOpJournal);

    const part = pkg.requirePart(slide.partUri);
    const withAdvancedLine = new TextDecoder().decode(part.bytes).replace(
      '<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="75000"/>' +
      '</a:srgbClr></a:solidFill><a:prstDash val="dashDot"/></a:ln>',
      '<a:gradFill><a:gsLst/></a:gradFill><a:custDash><a:ds d="1" sp="1"/>' +
      '</a:custDash><a:round/><a:headEnd type="triangle"/>' +
      '<a:tailEnd type="arrow"/><a:extLst><a:ext uri="urn:keep">' +
      '<x:lineKeep xmlns:x="urn:test"/></a:ext></a:extLst></a:ln>',
    );
    expect(withAdvancedLine).not.toEqual(new TextDecoder().decode(part.bytes));
    pkg.setPart(slide.partUri, withAdvancedLine, part.contentType);
    expect(shape.line).toBeUndefined();

    shape.line = {
      kind: 'line',
      color: { kind: 'scheme', value: 'accent3' },
      transparency: 40,
      width: 0,
      dash: 'sysDot',
    };
    expect(shape.line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent3' },
      transparency: 40,
      width: 0,
      dash: 'sysDot',
    });
    let xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml).toContain('<a:round/><a:headEnd type="triangle"/>');
    expect(xml).toContain('<a:tailEnd type="arrow"/>');
    expect(xml).toContain('<x:lineKeep xmlns:x="urn:test"/>');

    shape.line = { kind: 'none' };
    expect(shape.line).toEqual({ kind: 'none' });
    shape.line = undefined;
    expect(shape.line).toBeUndefined();
    xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml).toContain(
      '<a:ln><a:round/><a:headEnd type="triangle"/><a:tailEnd type="arrow"/>',
    );
    expect(xml).toContain('<x:lineKeep xmlns:x="urn:test"/>');
    expect(slide.shapes[0]).toBe(shape);
    expect(slide.shapes[1]).toBe(text);

    const invalidBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const invalidJournal = [...pkg.mutations];
    expect(() => {
      shape.line = {
        kind: 'line',
        color: { kind: 'srgb', value: 'FFFFFF' },
        width: 1_585,
      };
    }).toThrow(RangeError);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(invalidBytes);
    expect(pkg.mutations).toEqual(invalidJournal);
  });

  it('creates and edits strict text shape preset geometry through the shared live model', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const omittedSlide = model.addSlide();
    const undefinedSlide = model.addSlide();
    const omitted = omittedSlide.addText('Same geometry');
    const explicitUndefined = undefinedSlide.addText('Same geometry', {
      shape: undefined,
    } as never);
    expect(omitted.presetType).toBe('rect');
    expect(explicitUndefined.presetType).toBe('rect');
    expect(pkg.requirePart(undefinedSlide.partUri).bytes).toEqual(
      pkg.requirePart(omittedSlide.partUri).bytes,
    );

    const slide = model.addSlide();
    const plain = slide.addText('Ellipse text', { shape: 'ellipse' });
    const rich = slide.addRichText([{
      runs: [
        { text: 'Styled', style: { hyperlink: { url: 'https://run.example' } } },
        { text: ' geometry' },
      ],
    }], {
      name: 'Combined shaped text',
      shape: 'blockArc',
      x: inches(1),
      y: inches(2),
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

    expect(plain.presetType).toBe('ellipse');
    expect(rich.presetType).toBe('blockArc');
    expect(rich.text).toBe('Styled geometry');
    expect(rich.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent2' },
    });
    expect(rich.line).toMatchObject({ kind: 'line', color: { kind: 'srgb', value: '123ABC' } });
    expect(rich.arrows).toEqual({ begin: 'oval', end: 'triangle' });
    expect(rich.shadow).toMatchObject({ kind: 'outer', opacity: 0.5 });
    expect(rich.hyperlink).toEqual({ url: 'https://shape.example', tooltip: 'Shape' });
    expect(rich.textMargins).toEqual({ top: 0, right: 0, bottom: 0, left: 0 });
    expect(rich.verticalAlignment).toBe('bottom');
    expect(rich.textDirection).toBe('vert');
    expect(rich.textFit).toBe('shrink');
    expect(rich.textWrap).toBe(false);

    rich.adjustments = [
      { name: 'adj1', value: 16_200_000 },
      { name: 'adj2', value: 0 },
    ];
    const beforeSameType = packageSnapshot(pkg);
    rich.presetType = 'blockArc';
    expect(packageSnapshot(pkg)).toEqual(beforeSameType);
    expect(rich.adjustments).toEqual([
      { name: 'adj1', value: 16_200_000 },
      { name: 'adj2', value: 0 },
    ]);

    rich.presetType = 'hexagon';
    expect(rich.presetType).toBe('hexagon');
    expect(rich.adjustments).toEqual([]);
    expect(rich.text).toBe('Styled geometry');
    expect(rich.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent2' },
    });
    expect(rich.hyperlink).toEqual({ url: 'https://shape.example', tooltip: 'Shape' });
    expect(slide.shapes[1]).toBe(rich);

    rich.customGeometry = customTriangleGeometry;
    expect(rich.presetType).toBeUndefined();
    expect(rich.customGeometry).toEqual(customTriangleGeometry);
    expect(rich.text).toBe('Styled geometry');
    rich.presetType = 'star5';
    expect(rich.presetType).toBe('star5');
    expect(rich.customGeometry).toBeUndefined();

    const catalogSlide = model.addSlide();
    const catalogShapes = PRESET_SHAPE_TYPES.map((shape, index) => catalogSlide.addText(
      `Geometry ${index}`,
      { name: `text_geometry_${index}`, shape },
    ));
    expect(catalogShapes.map(({ presetType }) => presetType)).toEqual(PRESET_SHAPE_TYPES);
    expect(catalogSlide.shapes).toEqual(catalogShapes);

    const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
    const duplicateRich = duplicate.shapes.find(
      ({ name }) => name === 'Combined shaped text',
    ) as ShapeModel;
    expect(duplicateRich.presetType).toBe('star5');
    rich.presetType = 'diamond';
    expect(duplicateRich.presetType).toBe('star5');
    expect(rich.presetType).toBe('diamond');

    const beforeRollback = packageSnapshot(pkg);
    expect(() => pkg.transaction(() => {
      rich.presetType = 'triangle';
      throw new Error('restore shaped text geometry');
    })).toThrow('restore shaped text geometry');
    expect(packageSnapshot(pkg)).toEqual(beforeRollback);
    expect(rich.presetType).toBe('diamond');

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!;
    const reopenedRich = reopenedSlide.shapes.find(
      ({ name }) => name === 'Combined shaped text',
    ) as ShapeModel;
    expect(reopenedRich.presetType).toBe('diamond');
    expect(reopenedRich.text).toBe('Styled geometry');
    expect(reopenedRich.hyperlink).toEqual({ url: 'https://shape.example', tooltip: 'Shape' });
  }, 20_000);

  it('rejects invalid and malformed text shape preset geometry without mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const existing = slide.addText('Existing text');
    let accessorCalls = 0;
    const accessor = Object.defineProperty({}, 'shape', {
      enumerable: true,
      get() {
        accessorCalls += 1;
        throw new Error('shape getter must not run');
      },
    });
    const invalid = [
      { shape: '' },
      { shape: 'folderCorner' },
      { shape: 'custGeom' },
      { shape: 'unknown' },
      { shape: 1 },
      { shape: false },
      { shape: null },
      { shape: {} },
      { shape: Symbol('shape') },
      accessor,
    ];

    for (const options of invalid) {
      const before = packageSnapshot(pkg);
      const shapes = slide.shapes;
      expect(() => slide.addText('Invalid geometry', options as never)).toThrow(TypeError);
      expect(packageSnapshot(pkg)).toEqual(before);
      expect(slide.shapes).toEqual(shapes);
      expect(slide.shapes[0]).toBe(existing);
    }
    expect(accessorCalls).toBe(0);

    const inherited = Object.create({ shape: 'ellipse' }) as Record<string, unknown>;
    inherited.name = 'Inherited geometry';
    const inheritedShape = slide.addText('Inherited is ignored', inherited as never);
    expect(inheritedShape.presetType).toBe('rect');

    const malformed = slide.addText('Malformed geometry', { shape: 'ellipse' });
    const part = pkg.requirePart(slide.partUri);
    const malformedXml = new TextDecoder().decode(part.bytes).replace(
      '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>',
      '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>' +
      '<a:prstGeom prst="star5"><a:avLst/></a:prstGeom>',
    );
    pkg.setPart(slide.partUri, malformedXml, part.contentType);
    expect(malformed.presetType).toBeUndefined();
    const malformedBefore = packageSnapshot(pkg);
    expect(() => {
      malformed.presetType = 'diamond';
    }).toThrow(ModelParseError);
    expect(packageSnapshot(pkg)).toEqual(malformedBefore);
  });

  it('creates and edits text shape rectangle radius through direct adjustments', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const omittedSlide = model.addSlide();
    const undefinedSlide = model.addSlide();
    const omitted = omittedSlide.addText('Same radius', { shape: 'roundRect' });
    const explicitUndefined = undefinedSlide.addText('Same radius', {
      shape: 'roundRect',
      rectRadius: undefined,
    } as never);
    expect(omitted.adjustments).toEqual([]);
    expect(explicitUndefined.adjustments).toEqual([]);
    expect(pkg.requirePart(undefinedSlide.partUri).bytes).toEqual(
      pkg.requirePart(omittedSlide.partUri).bytes,
    );

    const slide = model.addSlide();
    const zero = slide.addText('Explicit zero radius', {
      shape: 'roundRect',
      rectRadius: inches(0),
      width: inches(2),
      height: inches(1),
    });
    const negativeZero = slide.addText('Negative zero radius', {
      shape: 'roundRect',
      rectRadius: -0,
    } as never);
    const twoByOne = slide.addText('Two by one', {
      shape: 'roundRect',
      rectRadius: inches(0.5),
      width: inches(2),
      height: inches(1),
    });
    const fourByTwo = slide.addText('Four by two', {
      name: 'Four by two radius',
      shape: 'roundRect',
      rectRadius: inches(0.5),
      width: inches(4),
      height: inches(2),
    });
    const boundary = slide.addText('Boundary radius', {
      shape: 'roundRect',
      rectRadius: inches(1),
      width: inches(2),
      height: inches(1),
    });
    const portrait = slide.addText('Portrait radius', {
      shape: 'roundRect',
      rectRadius: inches(0.25),
      width: inches(1),
      height: inches(2),
    });
    const fractional = slide.addText('Fractional EMU radius', {
      shape: 'roundRect',
      rectRadius: 1.4,
      width: 10.4,
      height: 20.4,
    } as never);
    const rich = slide.addRichText([{
      runs: [
        { text: 'Rounded', style: { hyperlink: { url: 'https://run.example' } } },
        { text: ' rich text' },
      ],
    }], {
      name: 'Combined rounded text',
      shape: 'roundRect',
      rectRadius: inches(0.5),
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
    const placeholder = slide.addPlaceholder('Rounded prompt', {
      name: 'rounded_prompt',
      type: 'title',
      shape: 'roundRect',
      rectRadius: inches(0.25),
      width: inches(2),
      height: inches(1),
    });

    expect(zero.adjustments).toEqual([{ name: 'adj', value: 0 }]);
    expect(negativeZero.adjustments).toEqual([{ name: 'adj', value: 0 }]);
    expect(Object.is(negativeZero.adjustments?.[0]?.value, -0)).toBe(false);
    expect(twoByOne.adjustments).toEqual([{ name: 'adj', value: 50_000 }]);
    expect(fourByTwo.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    expect(boundary.adjustments).toEqual([{ name: 'adj', value: 100_000 }]);
    expect(portrait.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    expect(fractional.adjustments).toEqual([{ name: 'adj', value: 10_000 }]);
    expect(rich.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    expect(placeholder.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    expect(rich.text).toBe('Rounded rich text');
    expect(rich.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent2' },
    });
    expect(rich.arrows).toEqual({ begin: 'oval', end: 'triangle' });
    expect(rich.shadow).toMatchObject({ kind: 'outer', opacity: 0.5 });
    expect(rich.hyperlink).toEqual({ url: 'https://shape.example', tooltip: 'Shape' });
    expect(rich.textMargins).toEqual({ top: 0, right: 0, bottom: 0, left: 0 });
    expect(rich.verticalAlignment).toBe('bottom');
    expect(rich.textDirection).toBe('vert');
    expect(rich.textFit).toBe('shrink');
    expect(rich.textWrap).toBe(false);

    const snapshot = twoByOne.adjustments!;
    expect(Object.isFrozen(snapshot)).toBe(true);
    expect(Object.isFrozen(snapshot[0])).toBe(true);
    const beforeNoOp = packageSnapshot(pkg);
    twoByOne.adjustments = [{ name: 'adj', value: 50_000 }];
    expect(packageSnapshot(pkg)).toEqual(beforeNoOp);
    twoByOne.setTransform({ width: inches(4), height: inches(2) });
    expect(twoByOne.adjustments).toEqual([{ name: 'adj', value: 50_000 }]);
    twoByOne.fill = { kind: 'solid', color: { kind: 'srgb', value: 'ABCDEF' } };
    twoByOne.line = {
      kind: 'line',
      color: { kind: 'scheme', value: 'accent3' },
      width: 1,
      dash: 'solid',
    };
    twoByOne.arrows = { begin: 'diamond', end: 'stealth' };
    twoByOne.shadow = { kind: 'inner', opacity: 0.25 };
    expect(twoByOne.adjustments).toEqual([{ name: 'adj', value: 50_000 }]);
    twoByOne.adjustments = [{ name: 'adj', value: 12_500 }];
    expect(twoByOne.adjustments).toEqual([{ name: 'adj', value: 12_500 }]);
    twoByOne.adjustments = [];
    expect(twoByOne.adjustments).toEqual([]);

    const beforeSamePreset = packageSnapshot(pkg);
    rich.presetType = 'roundRect';
    expect(packageSnapshot(pkg)).toEqual(beforeSamePreset);
    expect(rich.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);

    const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
    const duplicateRich = duplicate.shapes.find(
      ({ name }) => name === 'Combined rounded text',
    ) as ShapeModel;
    expect(duplicateRich.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
    model.moveSlide(model.slides.indexOf(duplicate), 0);
    expect(duplicateRich.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);

    const beforeRollback = packageSnapshot(pkg);
    expect(() => pkg.transaction(() => {
      rich.adjustments = [{ name: 'adj', value: 75_000 }];
      rich.setTransform({ width: inches(8), height: inches(4) });
      throw new Error('restore rounded text radius');
    })).toThrow('restore rounded text radius');
    expect(packageSnapshot(pkg)).toEqual(beforeRollback);
    expect(rich.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);

    rich.presetType = 'ellipse';
    expect(rich.adjustments).toEqual([]);
    expect(rich.text).toBe('Rounded rich text');
    expect(rich.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent2' },
    });

    const xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml).toContain(
      '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 0"/>' +
      '</a:avLst></a:prstGeom>',
    );
    expect(xml).toContain(
      '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/>' +
      '</a:avLst></a:prstGeom>',
    );
    expect(xml).toContain(
      '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 100000"/>' +
      '</a:avLst></a:prstGeom>',
    );

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!;
    const reopenedFourByTwo = reopenedSlide.shapes.find(
      ({ name }) => name === 'Four by two radius',
    ) as ShapeModel;
    expect(reopenedFourByTwo.adjustments).toEqual([{ name: 'adj', value: 25_000 }]);
  });

  it('rejects invalid text shape rectangle radius without mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const existing = slide.addText('Existing text');
    let accessorCalls = 0;
    const accessor = Object.defineProperty({ shape: 'roundRect' }, 'rectRadius', {
      enumerable: true,
      get() {
        accessorCalls += 1;
        throw new Error('rectRadius getter must not run');
      },
    });
    const invalid = [
      { options: { rectRadius: inches(0.5) }, error: TypeError },
      { options: { shape: 'rect', rectRadius: inches(0.5) }, error: TypeError },
      { options: { shape: 'ellipse', rectRadius: inches(0.5) }, error: TypeError },
      { options: { shape: 'roundRect', rectRadius: -0.1 }, error: RangeError },
      { options: { shape: 'roundRect', rectRadius: inches(1) + 0.1 }, error: RangeError },
      { options: { shape: 'roundRect', rectRadius: Number.NaN }, error: TypeError },
      {
        options: { shape: 'roundRect', rectRadius: Number.POSITIVE_INFINITY },
        error: TypeError,
      },
      {
        options: { shape: 'roundRect', rectRadius: Number.NEGATIVE_INFINITY },
        error: TypeError,
      },
      { options: { shape: 'roundRect', rectRadius: '0.5' }, error: TypeError },
      { options: { shape: 'roundRect', rectRadius: false }, error: TypeError },
      { options: { shape: 'roundRect', rectRadius: null }, error: TypeError },
      { options: { shape: 'roundRect', rectRadius: {} }, error: TypeError },
      { options: { shape: 'roundRect', rectRadius: Symbol('radius') }, error: TypeError },
      { options: accessor, error: TypeError },
    ];

    for (const { options, error } of invalid) {
      const before = packageSnapshot(pkg);
      const shapes = slide.shapes;
      expect(() => slide.addText('Invalid radius', options as never)).toThrow(error);
      expect(packageSnapshot(pkg)).toEqual(before);
      expect(slide.shapes).toEqual(shapes);
      expect(slide.shapes[0]).toBe(existing);
    }
    expect(accessorCalls).toBe(0);

    const inherited = Object.create({ rectRadius: inches(0.5) }) as Record<string, unknown>;
    inherited.name = 'Inherited radius';
    inherited.shape = 'roundRect';
    const inheritedShape = slide.addText('Inherited is ignored', inherited as never);
    expect(inheritedShape.presetType).toBe('roundRect');
    expect(inheritedShape.adjustments).toEqual([]);
  });

  it('creates and edits strict text box state independently from text shape content', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const omittedSlide = model.addSlide();
    const undefinedSlide = model.addSlide();
    const omitted = omittedSlide.addText('Same text box state');
    const explicitUndefined = undefinedSlide.addText('Same text box state', {
      isTextBox: undefined,
    } as never);
    expect(omitted.isTextBox).toBe(false);
    expect(explicitUndefined.isTextBox).toBe(false);
    expect(pkg.requirePart(undefinedSlide.partUri).bytes).toEqual(
      pkg.requirePart(omittedSlide.partUri).bytes,
    );

    const slide = model.addSlide();
    const explicitFalse = slide.addText('Shape text', {
      name: 'Explicit false text box',
      isTextBox: false,
    });
    const explicitTrue = slide.addText('Text box', {
      name: 'Explicit true text box',
      isTextBox: true,
    });
    const rich = slide.addRichText([{
      runs: [
        { text: 'Rich', style: { hyperlink: { url: 'https://run.example' } } },
        { text: ' text box' },
      ],
    }], {
      name: 'Combined text box state',
      isTextBox: true,
      shape: 'roundRect',
      rectRadius: inches(0.25),
      x: inches(1),
      y: inches(2),
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
    const placeholder = slide.addPlaceholder('Text box prompt', {
      name: 'text_box_prompt',
      type: 'title',
      isTextBox: true,
    });

    expect(explicitFalse.isTextBox).toBe(false);
    expect(explicitTrue.isTextBox).toBe(true);
    expect(rich.isTextBox).toBe(true);
    expect(placeholder.isTextBox).toBe(true);
    const createdXml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(createdXml).toContain(
      '<p:cNvPr id="2" name="Explicit false text box"/><p:cNvSpPr/>',
    );
    expect(createdXml).toContain(
      '<p:cNvPr id="3" name="Explicit true text box"/><p:cNvSpPr txBox="1"/>',
    );

    let beforeNoOp = packageSnapshot(pkg);
    explicitFalse.isTextBox = false;
    expect(packageSnapshot(pkg)).toEqual(beforeNoOp);
    beforeNoOp = packageSnapshot(pkg);
    explicitTrue.isTextBox = true;
    expect(packageSnapshot(pkg)).toEqual(beforeNoOp);

    explicitFalse.isTextBox = true;
    expect(explicitFalse.isTextBox).toBe(true);
    explicitTrue.isTextBox = false;
    expect(explicitTrue.isTextBox).toBe(false);
    expect(slide.shapes[0]).toBe(explicitFalse);
    expect(slide.shapes[1]).toBe(explicitTrue);

    rich.presetType = 'hexagon';
    rich.adjustments = [{ name: 'adj', value: 12_500 }];
    rich.setTransform({ width: inches(5) });
    rich.fill = { kind: 'solid', color: { kind: 'srgb', value: 'ABCDEF' } };
    rich.line = { kind: 'none' };
    rich.arrows = { begin: 'diamond', end: 'stealth' };
    rich.shadow = { kind: 'inner', opacity: 0.25 };
    rich.textMargins = 0;
    rich.verticalAlignment = 'top';
    rich.textDirection = 'vert270';
    rich.textFit = 'resize';
    rich.textWrap = true;
    rich.text = 'Edited rich text box';
    expect(rich.isTextBox).toBe(true);

    const beforeToggle = {
      presetType: rich.presetType,
      adjustments: rich.adjustments,
      transform: rich.transform,
      fill: rich.fill,
      line: rich.line,
      arrows: rich.arrows,
      shadow: rich.shadow,
      hyperlink: rich.hyperlink,
      margins: rich.textMargins,
      verticalAlignment: rich.verticalAlignment,
      direction: rich.textDirection,
      fit: rich.textFit,
      wrap: rich.textWrap,
      text: rich.text,
      placeholder: rich.placeholder,
    };
    rich.isTextBox = false;
    expect(rich.isTextBox).toBe(false);
    expect({
      presetType: rich.presetType,
      adjustments: rich.adjustments,
      transform: rich.transform,
      fill: rich.fill,
      line: rich.line,
      arrows: rich.arrows,
      shadow: rich.shadow,
      hyperlink: rich.hyperlink,
      margins: rich.textMargins,
      verticalAlignment: rich.verticalAlignment,
      direction: rich.textDirection,
      fit: rich.textFit,
      wrap: rich.textWrap,
      text: rich.text,
      placeholder: rich.placeholder,
    }).toEqual(beforeToggle);
    rich.isTextBox = true;

    const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
    const duplicateRich = duplicate.shapes.find(
      ({ name }) => name === 'Combined text box state',
    ) as ShapeModel;
    expect(duplicateRich.isTextBox).toBe(true);
    duplicateRich.isTextBox = false;
    expect(duplicateRich.isTextBox).toBe(false);
    expect(rich.isTextBox).toBe(true);
    model.moveSlide(model.slides.indexOf(duplicate), 0);
    expect(duplicateRich.isTextBox).toBe(false);

    const beforeRollback = packageSnapshot(pkg);
    expect(() => pkg.transaction(() => {
      rich.isTextBox = false;
      rich.presetType = 'triangle';
      throw new Error('restore text box state');
    })).toThrow('restore text box state');
    expect(packageSnapshot(pkg)).toEqual(beforeRollback);
    expect(rich.isTextBox).toBe(true);
    expect(rich.presetType).toBe('hexagon');

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!;
    const reopenedRich = reopenedSlide.shapes.find(
      ({ name }) => name === 'Combined text box state',
    ) as ShapeModel;
    expect(reopenedRich.isTextBox).toBe(true);
    expect(reopenedRich.text).toBe('Edited rich text box');
    expect(reopenedRich.presetType).toBe('hexagon');
    expect(reopenedRich.adjustments).toEqual([{ name: 'adj', value: 12_500 }]);
  });

  it('rejects invalid text box creation and editing and repairs single malformed tokens', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const existing = slide.addText('Existing text box', { isTextBox: true });
    let accessorCalls = 0;
    const accessor = Object.defineProperty({}, 'isTextBox', {
      enumerable: true,
      get() {
        accessorCalls += 1;
        throw new Error('isTextBox getter must not run');
      },
    });
    const invalid = [
      '',
      'true',
      0,
      1,
      null,
      {},
      [],
      new Boolean(true),
      () => true,
      Symbol('isTextBox'),
    ];
    for (const value of invalid) {
      const before = packageSnapshot(pkg);
      const shapes = slide.shapes;
      expect(() => slide.addText('Invalid text box', { isTextBox: value } as never))
        .toThrow(TypeError);
      expect(packageSnapshot(pkg)).toEqual(before);
      expect(slide.shapes).toEqual(shapes);
      expect(slide.shapes[0]).toBe(existing);
    }
    const beforeAccessor = packageSnapshot(pkg);
    expect(() => slide.addText('Accessor text box', accessor as never)).toThrow(TypeError);
    expect(accessorCalls).toBe(0);
    expect(packageSnapshot(pkg)).toEqual(beforeAccessor);

    const inherited = Object.create({ isTextBox: true }) as Record<string, unknown>;
    inherited.name = 'Inherited text box';
    const inheritedShape = slide.addText('Inherited is ignored', inherited as never);
    expect(inheritedShape.isTextBox).toBe(false);

    for (const value of invalid) {
      const before = packageSnapshot(pkg);
      expect(() => {
        (existing as unknown as { isTextBox: boolean }).isTextBox = value as never;
      }).toThrow(TypeError);
      expect(packageSnapshot(pkg)).toEqual(before);
      expect(existing.isTextBox).toBe(true);
    }

    const malformed = slide.addText('Malformed text box', {
      name: 'Malformed text box',
      isTextBox: true,
    });
    let part = pkg.requirePart(slide.partUri);
    let source = new TextDecoder().decode(part.bytes).replace(
      `<p:cNvPr id="${malformed.id}" name="Malformed text box"/><p:cNvSpPr txBox="1"/>`,
      `<p:cNvPr id="${malformed.id}" name="Malformed text box"/>`
        + '<p:cNvSpPr txBox="maybe"/>',
    );
    pkg.setPart(slide.partUri, source, part.contentType);
    expect(malformed.isTextBox).toBeUndefined();
    malformed.isTextBox = true;
    expect(malformed.isTextBox).toBe(true);
    expect(new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes)).toContain(
      `<p:cNvPr id="${malformed.id}" name="Malformed text box"/>`
        + '<p:cNvSpPr txBox="1"/>',
    );

    part = pkg.requirePart(slide.partUri);
    source = new TextDecoder().decode(part.bytes).replace(
      `<p:cNvPr id="${malformed.id}" name="Malformed text box"/><p:cNvSpPr txBox="1"/>`,
      `<p:cNvPr id="${malformed.id}" name="Malformed text box"/>`
        + '<p:cNvSpPr txBox="1" txBox="0"/>',
    );
    pkg.setPart(slide.partUri, source, part.contentType);
    expect(malformed.isTextBox).toBeUndefined();
    const beforeAmbiguous = packageSnapshot(pkg);
    expect(() => {
      malformed.isTextBox = false;
    }).toThrow(ModelParseError);
    expect(packageSnapshot(pkg)).toEqual(beforeAmbiguous);
  });

  it('round-trips text box state in all six presentation formats', async () => {
    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const pkg = await OpcPackage.open(await modelFixture(profile.presentationContentType));
      const model = new PresentationModel(pkg);
      const slide = model.addSlide();
      const becomesTrue = slide.addText(`${profile.format} true`, {
        name: `${profile.format} becomes true`,
      });
      const becomesFalse = slide.addText(`${profile.format} false`, {
        name: `${profile.format} becomes false`,
        isTextBox: true,
      });
      const rich = slide.addRichText([{ runs: [{ text: `${profile.format} rich` }] }], {
        name: `${profile.format} rich true`,
        isTextBox: true,
      });
      becomesTrue.isTextBox = true;
      becomesFalse.isTextBox = false;
      expect([becomesTrue.isTextBox, becomesFalse.isTextBox, rich.isTextBox])
        .toEqual([true, false, true]);

      const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
      expect(duplicate.shapes.map((shape) => (shape as ShapeModel).isTextBox))
        .toEqual([true, false, true]);
      const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
      expect(reopened.format).toBe(profile.format);
      const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!;
      expect(reopenedSlide.shapes.map((shape) => (shape as ShapeModel).isTextBox))
        .toEqual([true, false, true]);
    }
  });

  it('creates plain and rich text with strict direct fills', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const source = {
      kind: 'solid',
      color: { kind: 'srgb', value: '#ab12cd' },
      transparency: 25,
    } as ShapeFill;

    const omitted = slide.addText('Omitted text fill');
    const explicitUndefined = slide.addText('Undefined text fill', { fill: undefined as never });
    const none = slide.addText('Explicit text no-fill', { fill: { kind: 'none' } });
    const plain = slide.addText('Plain solid fill', { fill: source });
    const rich = slide.addRichText([{ runs: [{ text: 'Rich theme fill' }] }], {
      fill: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 0,
      },
    });
    const transparent = slide.addText('Fully transparent fill', {
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: '00AA00' },
        transparency: 100,
      },
    });

    expect(omitted.fill).toEqual({ kind: 'none' });
    expect(explicitUndefined.fill).toEqual({ kind: 'none' });
    expect(none.fill).toEqual({ kind: 'none' });
    expect(plain.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'AB12CD' },
      transparency: 25,
    });
    expect(rich.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 0,
    });
    expect(transparent.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '00AA00' },
      transparency: 100,
    });
    expect(slide.shapes.slice(-6)).toEqual([
      omitted,
      explicitUndefined,
      none,
      plain,
      rich,
      transparent,
    ]);

    (source as { color: { value: string }; transparency: number }).color.value = 'FFFFFF';
    (source as { color: { value: string }; transparency: number }).transparency = 50;
    expect(plain.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'AB12CD' },
      transparency: 25,
    });

    const xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml.match(
      /<a:prstGeom prst="rect"><a:avLst\/><\/a:prstGeom><a:noFill\/><a:ln><a:noFill\/><\/a:ln>/g,
    )).toHaveLength(3);
    expect(xml).toContain(
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:srgbClr val="AB12CD"><a:alpha val="75000"/>' +
      '</a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln>',
    );
    expect(xml).toContain(
      '<a:solidFill><a:schemeClr val="accent2"><a:alpha val="100000"/>' +
      '</a:schemeClr></a:solidFill>',
    );
    expect(xml).toContain(
      '<a:solidFill><a:srgbClr val="00AA00"><a:alpha val="0"/>' +
      '</a:srgbClr></a:solidFill>',
    );

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    expect(reopened.slides.at(-1)?.shapes.slice(-6).map((shape) =>
      shape instanceof ShapeModel ? shape.fill : undefined)).toEqual([
      { kind: 'none' },
      { kind: 'none' },
      { kind: 'none' },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'AB12CD' },
        transparency: 25,
      },
      {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 0,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: '00AA00' },
        transparency: 100,
      },
    ]);
  });

  it('rejects invalid text fill creation without mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const existing = slide.addText('Existing text');
    let getterCalls = 0;
    const accessor = Object.defineProperty({}, 'kind', {
      enumerable: true,
      get() {
        getterCalls += 1;
        return 'none';
      },
    });
    const symbol = Symbol('unsafe');
    const invalid = [
      null,
      [],
      new Date(),
      accessor,
      { kind: 'solid' },
      { kind: 'solid', color: { kind: 'srgb', value: 'GG0000' } },
      { kind: 'solid', color: { kind: 'scheme', value: 'unknown' } },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
        transparency: -1,
      },
      { kind: 'none', transparency: 1 },
      { kind: 'none', [symbol]: true },
      { type: 'none' },
    ];

    for (const [index, value] of invalid.entries()) {
      const before = packageSnapshot(pkg);
      const shapes = slide.shapes;
      expect(() => slide.addText('Invalid fill', { fill: value as never })).toThrow();
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid rich fill' }] }], {
        fill: value as never,
      })).toThrow();
      expect(packageSnapshot(pkg), `invalid fill ${index}`).toEqual(before);
      expect(slide.shapes).toEqual(shapes);
      expect(slide.shapes.at(-1)).toBe(existing);
    }
    expect(getterCalls).toBe(0);
  });

  it('creates plain and rich text with strict direct lines', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const source = {
      kind: 'line',
      color: { kind: 'srgb', value: '#ab12cd' },
      transparency: 25,
      width: 2.5,
      dash: 'dashDot',
    } as ShapeLine;
    const dashes: readonly ShapeLineDash[] = [
      'solid',
      'dash',
      'dashDot',
      'lgDash',
      'lgDashDot',
      'lgDashDotDot',
      'sysDash',
      'sysDot',
    ];

    const omitted = slide.addText('Omitted text line');
    const explicitUndefined = slide.addText('Undefined text line', { line: undefined } as never);
    const none = slide.addText('Explicit text no-line', { line: { kind: 'none' } });
    const plain = slide.addText('Plain solid line', { line: source });
    const rich = slide.addRichText([{ runs: [{ text: 'Rich theme line' }] }], {
      line: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 0,
        width: 0,
        dash: 'sysDot',
      },
    });
    const transparent = slide.addText('Fully transparent line', {
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: '00AA00' },
        transparency: 100,
      },
    });
    const dashed = dashes.map((dash) => slide.addText(`Text line ${dash}`, {
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
        dash,
      },
    }));

    expect(omitted.line).toEqual({ kind: 'none' });
    expect(explicitUndefined.line).toEqual({ kind: 'none' });
    expect(none.line).toEqual({ kind: 'none' });
    expect(plain.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: 'AB12CD' },
      transparency: 25,
      width: 2.5,
      dash: 'dashDot',
    });
    expect(rich.line).toEqual({
      kind: 'line',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 0,
      width: 0,
      dash: 'sysDot',
    });
    expect(transparent.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '00AA00' },
      transparency: 100,
      width: 1,
      dash: 'solid',
    });
    expect(dashed.map((shape) => shape.line)).toEqual(dashes.map((dash) => ({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      width: 1,
      dash,
    })));

    (source as { color: { value: string }; transparency: number; width: number; dash: string })
      .color.value = 'FFFFFF';
    (source as { transparency: number }).transparency = 50;
    (source as { width: number }).width = 9;
    (source as { dash: string }).dash = 'solid';
    expect(plain.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: 'AB12CD' },
      transparency: 25,
      width: 2.5,
      dash: 'dashDot',
    });

    let xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml.match(
      /<a:prstGeom prst="rect"><a:avLst\/><\/a:prstGeom><a:noFill\/><a:ln><a:noFill\/><\/a:ln>/g,
    )).toHaveLength(3);
    expect(xml).toContain(
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>' +
      '<a:ln w="31750"><a:solidFill><a:srgbClr val="AB12CD">' +
      '<a:alpha val="75000"/></a:srgbClr></a:solidFill>' +
      '<a:prstDash val="dashDot"/></a:ln>',
    );
    expect(xml).toContain(
      '<a:ln w="0"><a:solidFill><a:schemeClr val="accent2">' +
      '<a:alpha val="100000"/></a:schemeClr></a:solidFill>' +
      '<a:prstDash val="sysDot"/></a:ln>',
    );
    expect(xml).toContain(
      '<a:ln w="12700"><a:solidFill><a:srgbClr val="00AA00">' +
      '<a:alpha val="0"/></a:srgbClr></a:solidFill>' +
      '<a:prstDash val="solid"/></a:ln>',
    );
    for (const dash of dashes) {
      expect(xml).toContain(`<a:prstDash val="${dash}"/>`);
    }

    const beforeNoOp = pkg.requirePart(slide.partUri).bytes.slice();
    const beforeJournal = [...pkg.mutations];
    plain.line = {
      kind: 'line',
      color: { kind: 'srgb', value: 'AB12CD' },
      transparency: 25,
      width: 2.5,
      dash: 'dashDot',
    };
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeNoOp);
    expect(pkg.mutations).toEqual(beforeJournal);
    transparent.line = undefined;
    expect(transparent.line).toBeUndefined();
    xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml).toContain('<a:ln></a:ln></p:spPr>');
    expect(xml).toContain('Fully transparent line');

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedLines = reopened.slides.at(-1)?.shapes.slice(-dashes.length)
      .map((shape) => shape instanceof ShapeModel ? shape.line : undefined);
    expect(reopenedLines).toEqual(dashes.map((dash) => ({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      width: 1,
      dash,
    })));
  });

  it('rejects invalid text line creation without mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const existing = slide.addText('Existing text');
    let getterCalls = 0;
    const accessor = Object.defineProperty({}, 'kind', {
      enumerable: true,
      get() {
        getterCalls += 1;
        return 'none';
      },
    });
    const symbol = Symbol('unsafe');
    const invalid = [
      null,
      [],
      new Date(),
      accessor,
      { kind: 'line' },
      { kind: 'line', color: { kind: 'srgb', value: 'GG0000' } },
      { kind: 'line', color: { kind: 'scheme', value: 'unknown' } },
      {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        transparency: -1,
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        width: 1_585,
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        dash: 'dot',
      },
      { kind: 'none', width: 1 },
      { kind: 'none', [symbol]: true },
      { type: 'none' },
      { color: 'FF0000', dashType: 'dash' },
    ];

    for (const [index, value] of invalid.entries()) {
      const before = packageSnapshot(pkg);
      const shapes = slide.shapes;
      expect(() => slide.addText('Invalid line', { line: value } as never)).toThrow();
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid rich line' }] }], {
        line: value,
      } as never)).toThrow();
      expect(packageSnapshot(pkg), `invalid line ${index}`).toEqual(before);
      expect(slide.shapes).toEqual(shapes);
      expect(slide.shapes.at(-1)).toBe(existing);
    }
    expect(getterCalls).toBe(0);
  });

  it('creates plain and rich text with strict direct arrows', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const source: ShapeArrows = { begin: 'triangle', end: 'arrow' };
    const types: readonly ShapeArrowType[] = [
      'none',
      'arrow',
      'diamond',
      'oval',
      'stealth',
      'triangle',
    ];

    const omitted = slide.addText('Omitted text arrows');
    const explicitUndefined = slide.addText('Undefined text arrows', {
      arrows: undefined,
    } as never);
    const empty = slide.addText('Empty text arrows', { arrows: {} });
    const plain = slide.addText('Plain text arrows', { arrows: source });
    const rich = slide.addRichText([{ runs: [{ text: 'Rich text arrows' }] }], {
      arrows: { begin: 'none', end: 'stealth' },
    });
    const combined = slide.addText('Combined text arrows', {
      line: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 25,
        width: 2.5,
        dash: 'dashDot',
      },
      arrows: { begin: 'diamond', end: 'oval' },
    });
    const typed = types.map((type) => slide.addText(`Text arrows ${type}`, {
      arrows: { begin: type, end: type },
    }));

    expect(omitted.arrows).toBeUndefined();
    expect(explicitUndefined.arrows).toBeUndefined();
    expect(empty.arrows).toBeUndefined();
    expect(plain.arrows).toEqual({ begin: 'triangle', end: 'arrow' });
    expect(rich.arrows).toEqual({ begin: 'none', end: 'stealth' });
    expect(combined.arrows).toEqual({ begin: 'diamond', end: 'oval' });
    expect(typed.map((shape) => shape.arrows)).toEqual(types.map((type) => ({
      begin: type,
      end: type,
    })));
    expect(Object.isFrozen(plain.arrows)).toBe(true);
    expect(plain.arrows).not.toBe(plain.arrows);

    (source as { begin: ShapeArrowType }).begin = 'oval';
    (source as { end: ShapeArrowType }).end = 'diamond';
    expect(plain.arrows).toEqual({ begin: 'triangle', end: 'arrow' });

    let xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml.match(
      /<a:prstGeom prst="rect"><a:avLst\/><\/a:prstGeom><a:noFill\/><a:ln><a:noFill\/><\/a:ln>/g,
    )).toHaveLength(3);
    expect(xml).toContain(
      '<a:ln><a:noFill/><a:headEnd type="triangle"/>' +
      '<a:tailEnd type="arrow"/></a:ln>',
    );
    expect(xml).toContain(
      '<a:ln><a:noFill/><a:headEnd type="none"/>' +
      '<a:tailEnd type="stealth"/></a:ln>',
    );
    expect(xml).toContain(
      '<a:ln w="31750"><a:solidFill><a:schemeClr val="accent2">' +
      '<a:alpha val="75000"/></a:schemeClr></a:solidFill>' +
      '<a:prstDash val="dashDot"/><a:headEnd type="diamond"/>' +
      '<a:tailEnd type="oval"/></a:ln>',
    );
    for (const type of types) {
      expect(xml).toContain(
        `<a:headEnd type="${type}"/><a:tailEnd type="${type}"/>`,
      );
    }

    const beforeNoOp = pkg.requirePart(slide.partUri).bytes.slice();
    const beforeJournal = [...pkg.mutations];
    plain.arrows = { begin: 'triangle', end: 'arrow' };
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeNoOp);
    expect(pkg.mutations).toEqual(beforeJournal);

    rich.arrows = { end: 'oval' };
    expect(rich.arrows).toEqual({ end: 'oval' });
    expect(rich.line).toEqual({ kind: 'none' });
    plain.arrows = undefined;
    expect(plain.arrows).toBeUndefined();
    expect(plain.line).toEqual({ kind: 'none' });
    combined.line = undefined;
    expect(combined.line).toBeUndefined();
    expect(combined.arrows).toEqual({ begin: 'diamond', end: 'oval' });
    xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml).toContain(
      '<a:ln><a:headEnd type="diamond"/><a:tailEnd type="oval"/></a:ln>',
    );

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedTypes = reopened.slides.at(-1)?.shapes.slice(-types.length)
      .map((shape) => shape instanceof ShapeModel ? shape.arrows : undefined);
    expect(reopenedTypes).toEqual(types.map((type) => ({ begin: type, end: type })));
  });

  it('rejects invalid text arrow creation without mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const existing = slide.addText('Existing text');
    let getterCalls = 0;
    const accessor = Object.defineProperty({}, 'begin', {
      enumerable: true,
      get() {
        getterCalls += 1;
        return 'triangle';
      },
    });
    const symbol = Symbol('unsafe');
    const inherited = Object.create({ begin: 'triangle' });
    const invalid = [
      null,
      [],
      new Date(),
      accessor,
      inherited,
      { begin: '' },
      { end: 'bogus' },
      { beginArrowType: 'triangle' },
      { endArrowType: 'arrow' },
      { lineHead: 'triangle' },
      { lineTail: 'arrow' },
      { begin: 'triangle', extra: true },
      { begin: 'triangle', [symbol]: true },
    ];

    for (const [index, value] of invalid.entries()) {
      const before = packageSnapshot(pkg);
      const shapes = slide.shapes;
      expect(() => slide.addText('Invalid arrows', { arrows: value } as never)).toThrow();
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid rich arrows' }] }], {
        arrows: value,
      } as never)).toThrow();
      expect(packageSnapshot(pkg), `invalid arrows ${index}`).toEqual(before);
      expect(slide.shapes).toEqual(shapes);
      expect(slide.shapes.at(-1)).toBe(existing);
    }
    expect(getterCalls).toBe(0);
  });

  it('creates plain and rich text with strict direct shadows', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const omittedSlide = model.addSlide();
    const undefinedSlide = model.addSlide();
    omittedSlide.addText('Text shadow');
    undefinedSlide.addText('Text shadow', { shadow: undefined } as never);
    expect(pkg.requirePart(undefinedSlide.partUri).bytes).toEqual(
      pkg.requirePart(omittedSlide.partUri).bytes,
    );

    const slide = model.addSlide();
    const color: { kind: 'scheme'; value: 'accent2' | 'accent3' } = {
      kind: 'scheme',
      value: 'accent2',
    };
    const source: ShapeShadow = {
      kind: 'outer',
      color,
      opacity: 0.42,
      blur: 7.25,
      angle: 123.4,
      distance: 5.5,
      rotateWithShape: true,
    };
    const combined = slide.addText('Combined text shadow', {
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'DDEEFF' } },
      line: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
        width: 2,
        dash: 'dash',
      },
      arrows: { begin: 'triangle', end: 'arrow' },
      shadow: source,
    });
    const rich = slide.addRichText([{ runs: [{ text: 'Rich text shadow' }] }], {
      shadow: {
        kind: 'inner',
        color: { kind: 'srgb', value: '445566' },
        opacity: 0,
        blur: 0,
        angle: 0,
        distance: 0,
      },
    });
    const defaults = slide.addText('Default text shadow', {
      shadow: { kind: 'outer' },
    });
    color.value = 'accent3';

    expect(combined.shadow).toEqual({
      kind: 'outer',
      color: { kind: 'scheme', value: 'accent2' },
      opacity: 0.42,
      blur: 7.25,
      angle: 123.4,
      distance: 5.5,
      rotateWithShape: true,
    });
    expect(rich.shadow).toEqual({
      kind: 'inner',
      color: { kind: 'srgb', value: '445566' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
    });
    expect(defaults.shadow).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '000000' },
      opacity: 0.75,
      blur: 8,
      angle: 270,
      distance: 4,
      rotateWithShape: false,
    });
    const first = combined.shadow;
    const second = combined.shadow;
    expect(first).not.toBe(second);
    expect(first?.color).not.toBe(second?.color);
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.isFrozen(first?.color)).toBe(true);

    let xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml).toContain(
      '<a:solidFill><a:srgbClr val="DDEEFF"/></a:solidFill>' +
      '<a:ln w="25400"><a:solidFill><a:srgbClr val="112233"/></a:solidFill>' +
      '<a:prstDash val="dash"/><a:headEnd type="triangle"/>' +
      '<a:tailEnd type="arrow"/></a:ln><a:effectLst>' +
      '<a:outerShdw sx="100000" sy="100000" kx="0" ky="0" algn="bl" ' +
      'rotWithShape="1" blurRad="92075" dist="69850" dir="7404000">' +
      '<a:schemeClr val="accent2"><a:alpha val="42000"/></a:schemeClr>' +
      '</a:outerShdw></a:effectLst>',
    );
    expect(xml).toContain(
      '<a:innerShdw blurRad="0" dist="0" dir="0">' +
      '<a:srgbClr val="445566"><a:alpha val="0"/></a:srgbClr></a:innerShdw>',
    );

    const beforeNoOp = pkg.requirePart(slide.partUri).bytes.slice();
    const beforeJournal = [...pkg.mutations];
    combined.shadow = {
      kind: 'outer',
      color: { kind: 'scheme', value: 'accent2' },
      opacity: 0.42,
      blur: 7.25,
      angle: 123.4,
      distance: 5.5,
      rotateWithShape: true,
    };
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeNoOp);
    expect(pkg.mutations).toEqual(beforeJournal);

    combined.shadow = undefined;
    expect(combined.shadow).toBeUndefined();
    expect(combined.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'DDEEFF' },
    });
    expect(combined.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      width: 2,
      dash: 'dash',
    });
    expect(combined.arrows).toEqual({ begin: 'triangle', end: 'arrow' });
    combined.shadow = { kind: 'outer' };
    combined.fill = undefined;
    combined.line = undefined;
    combined.arrows = undefined;
    expect(combined.shadow).toEqual({
      kind: 'outer',
      color: { kind: 'srgb', value: '000000' },
      opacity: 0.75,
      blur: 8,
      angle: 270,
      distance: 4,
      rotateWithShape: false,
    });
    xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml).toContain('<a:ln></a:ln><a:effectLst><a:outerShdw');

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!;
    const reopenedRich = reopenedSlide.shapes.find(({ id }) => id === rich.id) as ShapeModel;
    expect(reopenedRich.shadow).toEqual({
      kind: 'inner',
      color: { kind: 'srgb', value: '445566' },
      opacity: 0,
      blur: 0,
      angle: 0,
      distance: 0,
    });
  });

  it('rejects invalid text shadow creation without mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const existing = slide.addText('Existing text');
    let getterCalls = 0;
    const accessor = Object.defineProperty({}, 'kind', {
      enumerable: true,
      get() {
        getterCalls += 1;
        return 'outer';
      },
    });
    const inherited = Object.create({ kind: 'outer' });
    const invalid = [
      null,
      [],
      new Date(),
      accessor,
      inherited,
      {},
      { kind: 'none' },
      { type: 'outer' },
      { kind: 'outer', offset: 4 },
      { kind: 'inner', rotateWithShape: true },
      { kind: 'outer', opacity: '0.5' },
      { kind: 'outer', opacity: Number.NaN },
      { kind: 'outer', blur: -1 },
      { kind: 'outer', blur: 101 },
      { kind: 'outer', angle: 360 },
      { kind: 'outer', distance: 201 },
      { kind: 'outer', color: { kind: 'srgb', value: 'FFF' } },
      { kind: 'outer', color: { kind: 'scheme', value: 'unknown' } },
      { kind: 'outer', [Symbol('unsafe')]: true },
    ];

    for (const [index, value] of invalid.entries()) {
      const before = packageSnapshot(pkg);
      const shapes = slide.shapes;
      expect(() => slide.addText('Invalid shadow', { shadow: value } as never)).toThrow();
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid rich shadow' }] }], {
        shadow: value,
      } as never)).toThrow();
      expect(packageSnapshot(pkg), `invalid shadow ${index}`).toEqual(before);
      expect(slide.shapes).toEqual(shapes);
      expect(slide.shapes.at(-1)).toBe(existing);
    }
    expect(getterCalls).toBe(0);
  });

  it('creates plain and rich text hyperlinks with one shared shape and run relationship', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const omittedSlide = model.addSlide();
    const undefinedSlide = model.addSlide();
    omittedSlide.addText('Text hyperlink');
    undefinedSlide.addText('Text hyperlink', { hyperlink: undefined } as never);
    expect(pkg.requirePart(undefinedSlide.partUri).bytes).toEqual(
      pkg.requirePart(omittedSlide.partUri).bytes,
    );
    expect(undefinedSlide.relationships).toEqual(omittedSlide.relationships);

    const source = model.addSlide();
    const target = model.addSlide();
    const urlInput: { url: string; tooltip?: string } = {
      url: 'https://example.com/path?a=1&b=2',
      tooltip: 'Visit <now> & learn',
    };
    const plain = source.addText('First\n\nSecond', {
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'DDEEFF' } },
      line: { kind: 'line', color: { kind: 'srgb', value: '112233' }, width: 2 },
      arrows: { begin: 'triangle', end: 'arrow' },
      shadow: { kind: 'outer' },
      hyperlink: urlInput,
    });
    const rich = source.addRichText([{
      runs: [
        { text: 'One' },
        {
          text: 'Two',
          style: {
            underline: false,
            color: { kind: 'scheme', value: 'accent2' },
          },
        },
        { text: '' },
      ],
    }], {
      hyperlink: {
        slide: model.slides.indexOf(target) + 1,
        tooltip: '',
      },
    });
    const empty = source.addText('', {
      hyperlink: { slide: model.slides.indexOf(source) + 1 },
    });
    urlInput.url = 'https://changed.example';
    urlInput.tooltip = 'Changed';

    expect(plain.hyperlink).toEqual({
      url: 'https://example.com/path?a=1&b=2',
      tooltip: 'Visit <now> & learn',
    });
    expect(Object.isFrozen(plain.hyperlink)).toBe(true);
    expect(rich.hyperlink).toEqual({
      slide: model.slides.indexOf(target) + 1,
      tooltip: '',
    });
    expect(empty.hyperlink).toEqual({ slide: model.slides.indexOf(source) + 1 });
    expect(plain.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'DDEEFF' },
    });
    expect(plain.line).toEqual({
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      width: 2,
      dash: 'solid',
    });
    expect(plain.arrows).toEqual({ begin: 'triangle', end: 'arrow' });
    expect(plain.shadow?.kind).toBe('outer');

    const urlRelationship = source.relationships.find(
      ({ type, target: relationshipTarget }) =>
        type === HYPERLINK_RELATIONSHIP
        && relationshipTarget === 'https://example.com/path?a=1&b=2',
    )!;
    const slideRelationship = source.relationships.find(
      ({ type, resolvedTarget }) =>
        type === SLIDE_RELATIONSHIP && resolvedTarget === target.partUri,
    )!;
    const selfRelationship = source.relationships.find(
      ({ type, resolvedTarget }) =>
        type === SLIDE_RELATIONSHIP && resolvedTarget === source.partUri,
    )!;
    expect(urlRelationship.targetMode).toBe('External');
    expect(slideRelationship.targetMode).toBe('Internal');
    expect(selfRelationship.targetMode).toBe('Internal');

    const xml = new TextDecoder().decode(pkg.requirePart(source.partUri).bytes);
    expect(xml.match(new RegExp(`r:id="${urlRelationship.id}"`, 'g'))).toHaveLength(3);
    expect(xml.match(new RegExp(`r:id="${slideRelationship.id}"`, 'g'))).toHaveLength(3);
    expect(xml.match(new RegExp(`r:id="${selfRelationship.id}"`, 'g'))).toHaveLength(1);
    expect(xml).toContain('tooltip="Visit &lt;now&gt; &amp; learn"');
    expect(xml.match(new RegExp(
      `r:id="${slideRelationship.id}" tooltip="" action="ppaction://hlinksldjump"`,
      'g',
    ))).toHaveLength(3);
    expect(xml.match(/u="sng"/g)?.length).toBeGreaterThanOrEqual(3);
    expect(xml).toContain('u="none"');

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSource = reopened.slides.find(({ partUri }) => partUri === source.partUri)!;
    expect((reopenedSource.shapes.find(({ id }) => id === plain.id) as ShapeModel).hyperlink)
      .toEqual({
        url: 'https://example.com/path?a=1&b=2',
        tooltip: 'Visit <now> & learn',
      });
    expect((reopenedSource.shapes.find(({ id }) => id === rich.id) as ShapeModel).hyperlink)
      .toEqual({ slide: reopened.slides.indexOf(
        reopened.slides.find(({ partUri }) => partUri === target.partUri)!,
      ) + 1, tooltip: '' });
  });

  it('creates independent rich text run hyperlink relationships', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const source = model.addSlide();
    const target = model.addSlide();
    const targetIndex = model.slides.indexOf(target) + 1;

    const shape = source.addRichText([{
      runs: [
        { text: 'Outer' },
        {
          text: 'Local one',
          style: { hyperlink: { url: 'https://local.example/path', tooltip: 'Local' } },
        },
        { text: 'Suppressed', style: { bold: true, hyperlink: false } },
        {
          text: 'Internal',
          style: { hyperlink: { slide: targetIndex, tooltip: '' }, underline: false },
        },
        {
          text: 'Local two',
          style: { hyperlink: { url: 'https://local.example/path', tooltip: 'Second' } },
        },
      ],
    }], {
      hyperlink: { url: 'https://outer.example/path', tooltip: 'Outer' },
    });
    const placeholder = source.addPlaceholder([{
      runs: [{ text: 'Prompt link', style: { hyperlink: { url: 'https://prompt.example' } } }],
    }], {
      name: 'run_link_prompt',
      type: 'body',
      index: 71,
    });

    expect(shape.hyperlink).toEqual({
      url: 'https://outer.example/path',
      tooltip: 'Outer',
    });
    expect(placeholder.hyperlink).toBeUndefined();
    const outer = source.relationships.find(
      ({ type, target: relationshipTarget }) =>
        type === HYPERLINK_RELATIONSHIP
        && relationshipTarget === 'https://outer.example/path',
    )!;
    const locals = source.relationships.filter(
      ({ type, target: relationshipTarget }) =>
        type === HYPERLINK_RELATIONSHIP
        && relationshipTarget === 'https://local.example/path',
    );
    const internal = source.relationships.find(
      ({ type, resolvedTarget }) =>
        type === SLIDE_RELATIONSHIP && resolvedTarget === target.partUri,
    )!;
    const prompt = source.relationships.find(
      ({ type, target: relationshipTarget }) =>
        type === HYPERLINK_RELATIONSHIP
        && relationshipTarget === 'https://prompt.example',
    )!;

    expect(locals).toHaveLength(2);
    expect(new Set(locals.map(({ id }) => id)).size).toBe(2);
    expect(internal.targetMode).toBe('Internal');
    expect(prompt.targetMode).toBe('External');

    const xml = new TextDecoder().decode(pkg.requirePart(source.partUri).bytes);
    expect(xml).toContain(`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`);
    expect(xml.match(new RegExp(`r:id="${outer.id}"`, 'g'))).toHaveLength(2);
    expect(xml.match(new RegExp(`r:id="${locals[0]!.id}"`, 'g'))).toHaveLength(1);
    expect(xml.match(new RegExp(`r:id="${locals[1]!.id}"`, 'g'))).toHaveLength(1);
    expect(xml.match(new RegExp(`r:id="${internal.id}"`, 'g'))).toHaveLength(1);
    expect(xml.match(new RegExp(`r:id="${prompt.id}"`, 'g'))).toHaveLength(1);
    expect(xml).toContain(`r:id="${internal.id}" tooltip="" action="ppaction://hlinksldjump"`);
    expect(xml).toMatch(/u="none"[^>]*>[\s\S]*?<a:hlinkClick[^>]*action="ppaction:\/\/hlinksldjump"/);
    expect(shape.richText[0]!.runs.map((run) => run.style?.hyperlink)).toEqual([
      { url: 'https://outer.example/path', tooltip: 'Outer' },
      { url: 'https://local.example/path', tooltip: 'Local' },
      undefined,
      { slide: targetIndex, tooltip: '' },
      { url: 'https://local.example/path', tooltip: 'Second' },
    ]);
    expect(placeholder.richText[0]!.runs[0]!.style?.hyperlink)
      .toEqual({ url: 'https://prompt.example' });
  });

  it('creates and edits canonical rich text line break paragraphs without changing shape state', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const shape = slide.addRichText([{
      align: 'center',
      rtl: true,
      marginLeft: 12,
      marginRight: 18,
      indent: 4,
      level: 2,
      spacing: { before: 3, after: 5, line: { kind: 'multiple', factor: 1.25 } },
      tabStops: [{ position: 1.5, alignment: 'right' }],
      runs: [
        { text: 'First', style: { bold: true }, breakLine: true },
        { text: '', breakLine: true },
        { text: 'Soft', softBreakBefore: true },
        { text: 'Final', breakLine: true },
      ],
    }], {
      name: 'Canonical breakLine shape',
      shape: 'roundRect',
      rectRadius: inches(0.2),
      isTextBox: true,
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'DDEEFF' } },
      line: { kind: 'line', color: { kind: 'srgb', value: '112233' }, width: 2 },
      arrows: { begin: 'triangle', end: 'arrow' },
      shadow: { kind: 'outer' },
      margin: [2, 3, 4, 5],
      valign: 'bottom',
      vert: 'horz',
      fit: 'shrink',
      wrap: false,
    });

    const properties = {
      align: 'center' as const,
      rtl: true,
      marginLeft: 12,
      marginRight: 18,
      indent: 4,
      level: 2,
      spacing: { before: 3, after: 5, line: { kind: 'multiple' as const, factor: 1.25 } },
      tabStops: [{ position: 1.5, alignment: 'right' as const }],
    };
    const defaultStyle = {
      fontFamily: '+mn-lt',
      lang: 'en-US',
      color: { kind: 'scheme' as const, value: 'tx1' },
    };
    expect(shape.richText).toEqual([
      { ...properties, runs: [{ text: 'First', style: { ...defaultStyle, bold: true } }] },
      { ...properties, runs: [] },
      {
        ...properties,
        runs: [
          { text: 'Soft', style: defaultStyle, softBreakBefore: true },
          { text: 'Final', style: defaultStyle },
        ],
      },
    ]);
    expect(shape.text).toBe('First\n\n\nSoftFinal');
    expect(shape.richText.flatMap(({ runs }) => runs).some((run) =>
      Object.hasOwn(run, 'breakLine'))).toBe(false);

    const xml = LosslessXmlDocument.parse(pkg.requirePart(slide.partUri).bytes);
    const element = xml.elements('sp').find((candidate) => {
      const properties = xml.descendants(candidate, 'cNvPr')[0];
      return properties && xml.attribute(properties, 'id')?.value === String(shape.id);
    })!;
    const textBody = xml.descendants(element, 'txBody')[0]!;
    expect(textBody.children.filter((child) =>
      child.type === 'element' && child.localName === 'p')).toHaveLength(3);

    const beforeNoOp = packageSnapshot(pkg);
    shape.richText = shape.richText;
    expect(packageSnapshot(pkg)).toEqual(beforeNoOp);

    const ownedState = {
      presetType: shape.presetType,
      adjustments: shape.adjustments,
      isTextBox: shape.isTextBox,
      fill: shape.fill,
      line: shape.line,
      arrows: shape.arrows,
      shadow: shape.shadow,
      textMargins: shape.textMargins,
      verticalAlignment: shape.verticalAlignment,
      textDirection: shape.textDirection,
      textFit: shape.textFit,
      textWrap: shape.textWrap,
    };
    shape.richText = [{
      align: 'right',
      runs: [
        { text: 'Left', breakLine: true },
        { text: 'Right' },
      ],
    }];
    expect(shape.richText).toEqual([
      { align: 'right', runs: [{ text: 'Left', style: defaultStyle }] },
      { align: 'right', runs: [{ text: 'Right', style: defaultStyle }] },
    ]);
    expect({
      presetType: shape.presetType,
      adjustments: shape.adjustments,
      isTextBox: shape.isTextBox,
      fill: shape.fill,
      line: shape.line,
      arrows: shape.arrows,
      shadow: shape.shadow,
      textMargins: shape.textMargins,
      verticalAlignment: shape.verticalAlignment,
      textDirection: shape.textDirection,
      textFit: shape.textFit,
      textWrap: shape.textWrap,
    }).toEqual(ownedState);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!;
    expect((reopenedSlide.shapes.find(({ id }) => id === shape.id) as ShapeModel).richText)
      .toEqual(shape.richText);
  });

  it('reindexes rich text line break hyperlinks through edits, moves, duplicate, and rollback', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const source = model.addSlide();
    const target = model.addSlide();
    const targetIndex = model.slides.indexOf(target) + 1;
    const shape = source.addRichText([{
      runs: [
        {
          text: 'URL',
          breakLine: true,
          style: { hyperlink: { url: 'https://break.example', tooltip: 'URL' } },
        },
        {
          text: 'Slide',
          breakLine: true,
          style: { hyperlink: { slide: targetIndex, tooltip: '' } },
        },
        { text: 'Plain' },
      ],
    }]);

    expect(shape.richText.map(({ runs }) => runs.map(({ text }) => text))).toEqual([
      ['URL'],
      ['Slide'],
      ['Plain'],
    ]);
    let urlRelationships = source.relationships.filter(
      ({ type, target: value }) =>
        type === HYPERLINK_RELATIONSHIP && value === 'https://break.example',
    );
    let slideRelationships = source.relationships.filter(
      ({ type, resolvedTarget }) =>
        type === SLIDE_RELATIONSHIP && resolvedTarget === target.partUri,
    );
    expect(urlRelationships).toHaveLength(1);
    expect(slideRelationships).toHaveLength(1);
    const originalSlideRelationshipId = slideRelationships[0]!.id;

    shape.richText = [{
      runs: [
        {
          text: 'URL moved',
          style: { hyperlink: { url: 'https://break.example', tooltip: 'URL' } },
        },
        {
          text: 'Slide moved',
          breakLine: true,
          style: { hyperlink: { slide: targetIndex, tooltip: '' } },
        },
        { text: 'Plain moved' },
      ],
    }];
    expect(shape.richText.map(({ runs }) => runs.map(({ text }) => text))).toEqual([
      ['URL moved', 'Slide moved'],
      ['Plain moved'],
    ]);
    urlRelationships = source.relationships.filter(
      ({ type, target: value }) =>
        type === HYPERLINK_RELATIONSHIP && value === 'https://break.example',
    );
    slideRelationships = source.relationships.filter(
      ({ type, resolvedTarget }) =>
        type === SLIDE_RELATIONSHIP && resolvedTarget === target.partUri,
    );
    expect(urlRelationships).toHaveLength(1);
    expect(slideRelationships).toHaveLength(1);
    expect(source.relationships.some(({ id }) => id === originalSlideRelationshipId)).toBe(false);

    const beforeRollback = packageSnapshot(pkg);
    const textBeforeRollback = shape.richText;
    expect(() => pkg.transaction(() => {
      shape.richText = [{
        runs: [
          {
            text: 'Rollback URL',
            breakLine: true,
            style: { hyperlink: { url: 'https://rollback-break.example' } },
          },
          { text: 'Rollback plain' },
        ],
      }];
      throw new Error('restore rich text line break edit');
    })).toThrow('restore rich text line break edit');
    expect(packageSnapshot(pkg)).toEqual(beforeRollback);
    expect(shape.richText).toEqual(textBeforeRollback);

    model.moveSlide(model.slides.indexOf(target), 0);
    const movedTargetIndex = model.slides.indexOf(target) + 1;
    expect(shape.richText[0]!.runs[1]!.style?.hyperlink)
      .toEqual({ slide: movedTargetIndex, tooltip: '' });
    const duplicate = model.duplicateSlide(model.slides.indexOf(source));
    const duplicateShape = duplicate.shapes.find(({ name }) => name === shape.name) as ShapeModel;
    expect(duplicateShape.richText).toEqual(shape.richText);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSource = reopened.slides.find(({ partUri }) => partUri === source.partUri)!;
    const reopenedShape = reopenedSource.shapes.find(({ id }) => id === shape.id) as ShapeModel;
    expect(reopenedShape.richText.map(({ runs }) => runs.map(({ text }) => text))).toEqual([
      ['URL moved', 'Slide moved'],
      ['Plain moved'],
    ]);
    expect(reopenedShape.richText[0]!.runs[1]!.style?.hyperlink).toEqual({
      slide: reopened.slides.findIndex(({ partUri }) => partUri === target.partUri) + 1,
      tooltip: '',
    });
  });

  it('rejects invalid rich text line breaks atomically without executing accessors', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const shape = slide.addRichText([{ runs: [{ text: 'Existing' }] }]);
    let accessorCalls = 0;
    const accessor = Object.defineProperty({ text: 'Accessor' }, 'breakLine', {
      enumerable: true,
      get() {
        accessorCalls += 1;
        return true;
      },
    });
    const invalidInputs = [
      [{
        runs: [
          {
            text: 'Prepared hyperlink',
            style: { hyperlink: { url: 'https://prepared-break.example' } },
          },
          { text: 'Invalid later', breakLine: 'yes' },
        ],
      }],
      [{ runs: [accessor] }],
    ];

    for (const [index, input] of invalidInputs.entries()) {
      let before = packageSnapshot(pkg);
      const shapes = slide.shapes;
      expect(() => slide.addRichText(input as never, {
        hyperlink: { url: 'https://outer-break.example' },
      })).toThrow();
      expect(packageSnapshot(pkg), `create breakLine invalid ${index}`).toEqual(before);
      expect(slide.shapes).toEqual(shapes);
      expect(slide.shapes.at(-1)).toBe(shape);

      before = packageSnapshot(pkg);
      const richText = shape.richText;
      expect(() => {
        shape.richText = input as never;
      }).toThrow();
      expect(packageSnapshot(pkg), `edit breakLine invalid ${index}`).toEqual(before);
      expect(shape.richText).toEqual(richText);
      expect(slide.shapes.at(-1)).toBe(shape);
    }
    expect(accessorCalls).toBe(0);
  });

  it('round-trips rich text line breaks in all six presentation formats', async () => {
    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const pkg = await OpcPackage.open(await modelFixture(profile.presentationContentType));
      const model = new PresentationModel(pkg);
      const slide = model.addSlide();
      const shape = slide.addRichText([{
        runs: [
          { text: `${profile.format} first`, breakLine: true },
          { text: '', breakLine: true },
          { text: `${profile.format} last`, softBreakBefore: true, breakLine: true },
        ],
      }], { name: `${profile.format} breakLine` });
      expect(shape.richText.map(({ runs }) => runs.map(({ text }) => text))).toEqual([
        [`${profile.format} first`],
        [],
        [`${profile.format} last`],
      ]);

      const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
      model.moveSlide(model.slides.indexOf(duplicate), 0);
      const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
      expect(reopened.format).toBe(profile.format);
      for (const partUri of [slide.partUri, duplicate.partUri]) {
        const reopenedSlide = reopened.slides.find((candidate) => candidate.partUri === partUri)!;
        const reopenedShape = reopenedSlide.shapes.find(
          ({ name }) => name === `${profile.format} breakLine`,
        ) as ShapeModel;
        expect(reopenedShape.richText.map(({ runs }) => runs.map(({ text }) => text))).toEqual([
          [`${profile.format} first`],
          [],
          [`${profile.format} last`],
        ]);
        expect(reopenedShape.richText[2]!.runs[0]!.softBreakBefore).toBe(true);
      }
    }
  });

  it('rejects invalid rich text run hyperlink creation without mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const source = model.addSlide();
    model.addSlide();
    const existing = source.addText('Existing text');
    const invalidValues = [
      [{
        runs: [
          { text: 'Prepared first', style: { hyperlink: { url: 'https://prepared.example' } } },
          { text: 'Invalid later', style: { hyperlink: { slide: model.slides.length + 1 } } },
        ],
      }],
      [{
        runs: [{ text: '', style: { hyperlink: { url: 'https://empty.example' } } }],
      }],
    ];

    for (const [index, value] of invalidValues.entries()) {
      let before = packageSnapshot(pkg);
      let shapes = source.shapes;
      expect(() => source.addRichText(value as never, {
        hyperlink: { url: 'https://outer.example' },
      })).toThrow();
      expect(packageSnapshot(pkg), `rich text invalid run hyperlink ${index}`).toEqual(before);
      expect(source.shapes).toEqual(shapes);
      expect(source.shapes.at(-1)).toBe(existing);

      before = packageSnapshot(pkg);
      shapes = source.shapes;
      expect(() => source.addPlaceholder(value as never, {
        name: `invalid_run_link_${index}`,
        type: 'body',
        index: 80 + index,
        hyperlink: { url: 'https://outer.example' },
      })).toThrow();
      expect(packageSnapshot(pkg), `placeholder invalid run hyperlink ${index}`).toEqual(before);
      expect(source.shapes).toEqual(shapes);
      expect(source.shapes.at(-1)).toBe(existing);
    }
  });

  it('keeps text run hyperlinks independent from whole-shape hyperlink edits', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const source = model.addSlide();
    const target = model.addSlide();
    const shape = source.addText('Linked text', {
      hyperlink: { url: 'https://runs.example', tooltip: 'Runs' },
    });
    const originalRelationship = source.relationships.find(
      ({ type }) => type === HYPERLINK_RELATIONSHIP,
    )!;

    const noOp = packageSnapshot(pkg);
    shape.hyperlink = { url: 'https://runs.example', tooltip: 'Runs' };
    expect(packageSnapshot(pkg)).toEqual(noOp);

    shape.hyperlink = { url: 'https://runs.example', tooltip: '' };
    expect(source.relationships.find(({ id }) => id === originalRelationship.id)?.target)
      .toBe('https://runs.example');
    let xml = new TextDecoder().decode(pkg.requirePart(source.partUri).bytes);
    expect(xml.match(new RegExp(`r:id="${originalRelationship.id}"`, 'g'))).toHaveLength(2);

    shape.hyperlink = { slide: model.slides.indexOf(target) + 1 };
    const targetRelationship = source.relationships.find(
      ({ type, resolvedTarget }) => type === SLIDE_RELATIONSHIP && resolvedTarget === target.partUri,
    )!;
    expect(targetRelationship.id).not.toBe(originalRelationship.id);
    xml = new TextDecoder().decode(pkg.requirePart(source.partUri).bytes);
    expect(xml.match(new RegExp(`r:id="${originalRelationship.id}"`, 'g'))).toHaveLength(1);
    expect(xml.match(new RegExp(`r:id="${targetRelationship.id}"`, 'g'))).toHaveLength(1);
    expect(shape.hyperlink).toEqual({ slide: model.slides.indexOf(target) + 1 });

    shape.hyperlink = undefined;
    expect(shape.hyperlink).toBeUndefined();
    expect(source.relationships.some(({ id }) => id === targetRelationship.id)).toBe(false);
    expect(source.relationships.some(({ id }) => id === originalRelationship.id)).toBe(true);
    xml = new TextDecoder().decode(pkg.requirePart(source.partUri).bytes);
    expect(xml.match(new RegExp(`r:id="${originalRelationship.id}"`, 'g'))).toHaveLength(1);
    expect(xml).toContain('<a:rPr');
    expect(xml).toContain('<a:hlinkClick');
  });

  it('edits rich text run hyperlinks with reuse, COW, GC, rollback, and reopen', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const source = model.addSlide();
    const firstTarget = model.addSlide();
    const secondTarget = model.addSlide();
    const shape = source.addRichText([{
      runs: [
        { text: 'Unique URL', style: { hyperlink: { url: 'https://one.example', tooltip: 'One' } } },
        {
          text: 'Unique slide',
          style: { hyperlink: { slide: model.slides.indexOf(firstTarget) + 1 } },
        },
        { text: 'Shared outer' },
      ],
    }], {
      hyperlink: { url: 'https://outer-edit.example', tooltip: 'Outer' },
    });
    const editRun = (
      paragraphs: readonly RichTextParagraph[],
      runIndex: number,
      hyperlink: Hyperlink | undefined,
    ): readonly RichTextParagraph[] => paragraphs.map((paragraph, paragraphIndex) => ({
      ...paragraph,
      runs: paragraph.runs.map((run, candidateIndex) => {
        if (paragraphIndex !== 0 || candidateIndex !== runIndex) return run;
        const { hyperlink: currentHyperlink, ...style } = run.style ?? {};
        void currentHyperlink;
        return {
          ...run,
          style: hyperlink === undefined ? style : { ...style, hyperlink },
        };
      }),
    }));

    const uniqueUrl = source.relationships.find(
      ({ type, target }) => type === HYPERLINK_RELATIONSHIP && target === 'https://one.example',
    )!;
    const uniqueSlide = source.relationships.find(
      ({ type, resolvedTarget }) =>
        type === SLIDE_RELATIONSHIP && resolvedTarget === firstTarget.partUri,
    )!;
    const outer = source.relationships.find(
      ({ type, target }) =>
        type === HYPERLINK_RELATIONSHIP && target === 'https://outer-edit.example',
    )!;

    const snapshot = shape.richText;
    const beforeNoOp = packageSnapshot(pkg);
    shape.richText = snapshot;
    expect(packageSnapshot(pkg)).toEqual(beforeNoOp);

    shape.richText = editRun(shape.richText, 0, {
      url: 'https://one.example',
      tooltip: 'Updated tooltip',
    });
    expect(source.relationships.find(({ id }) => id === uniqueUrl.id)?.target)
      .toBe('https://one.example');
    expect(new TextDecoder().decode(pkg.requirePart(source.partUri).bytes))
      .toContain(`r:id="${uniqueUrl.id}" tooltip="Updated tooltip"`);

    shape.richText = editRun(shape.richText, 0, {
      url: 'https://changed.example',
      tooltip: 'Changed target',
    });
    expect(source.relationships.find(({ id }) => id === uniqueUrl.id)).toMatchObject({
      target: 'https://changed.example',
      targetMode: 'External',
    });

    shape.richText = editRun(shape.richText, 2, {
      slide: model.slides.indexOf(secondTarget) + 1,
    });
    const copied = source.relationships.find(
      ({ type, resolvedTarget }) =>
        type === SLIDE_RELATIONSHIP && resolvedTarget === secondTarget.partUri,
    )!;
    expect(copied.id).not.toBe(outer.id);
    expect(source.relationships.find(({ id }) => id === outer.id)?.target)
      .toBe('https://outer-edit.example');
    expect(shape.hyperlink).toEqual({ url: 'https://outer-edit.example', tooltip: 'Outer' });

    shape.richText = editRun(shape.richText, 0, undefined);
    expect(source.relationships.some(({ id }) => id === uniqueUrl.id)).toBe(false);
    shape.richText = editRun(shape.richText, 1, undefined);
    expect(source.relationships.some(({ id }) => id === uniqueSlide.id)).toBe(false);
    shape.richText = editRun(shape.richText, 2, undefined);
    expect(source.relationships.some(({ id }) => id === copied.id)).toBe(false);
    expect(source.relationships.some(({ id }) => id === outer.id)).toBe(true);

    const beforeInvalid = packageSnapshot(pkg);
    expect(() => {
      shape.richText = editRun(editRun(shape.richText, 0, {
        url: 'https://prepared-edit.example',
      }), 2, { slide: model.slides.length + 1 });
    }).toThrow();
    expect(packageSnapshot(pkg)).toEqual(beforeInvalid);

    const beforeRollback = packageSnapshot(pkg);
    const richTextBeforeRollback = shape.richText;
    expect(() => pkg.transaction(() => {
      shape.richText = editRun(shape.richText, 1, {
        slide: model.slides.indexOf(firstTarget) + 1,
        tooltip: 'Rollback',
      });
      throw new Error('rollback rich text run hyperlink edit');
    })).toThrow('rollback rich text run hyperlink edit');
    expect(packageSnapshot(pkg)).toEqual(beforeRollback);
    expect(shape.richText).toEqual(richTextBeforeRollback);

    shape.richText = editRun(shape.richText, 0, {
      url: 'https://reopen.example',
      tooltip: '',
    });
    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSource = reopened.slides.find(({ partUri }) => partUri === source.partUri)!;
    const reopenedShape = reopenedSource.shapes.find(({ id }) => id === shape.id) as ShapeModel;
    expect(reopenedShape.richText[0]!.runs[0]!.style?.hyperlink)
      .toEqual({ url: 'https://reopen.example', tooltip: '' });
    expect(reopenedShape.hyperlink)
      .toEqual({ url: 'https://outer-edit.example', tooltip: 'Outer' });
  });

  it('rejects invalid text hyperlink creation and rolls valid creation back exactly', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    model.addSlide();
    const existing = slide.addText('Existing text');
    let getterCalls = 0;
    const accessor = Object.defineProperty({}, 'url', {
      enumerable: true,
      get() {
        getterCalls += 1;
        return 'https://example.com';
      },
    });
    const inherited = Object.create({ url: 'https://example.com' });
    const invalid = [
      null,
      false,
      [],
      new Date(),
      accessor,
      inherited,
      {},
      { url: '', slide: undefined },
      { url: 'https://example.com', slide: 2 },
      { url: 123 },
      { slide: 0 },
      { slide: -1 },
      { slide: 1.5 },
      { slide: Number.NaN },
      { slide: Number.POSITIVE_INFINITY },
      { slide: Number.MAX_SAFE_INTEGER + 1 },
      { slide: model.slides.length + 1 },
      { url: 'bad\u0000url' },
      { url: 'https://example.com', tooltip: 'bad\u0000tooltip' },
      { target: 'https://example.com' },
      { url: 'https://example.com', _rId: 'rId9' },
      { url: 'https://example.com', kind: 'url' },
      { url: 'https://example.com', [Symbol('unsafe')]: true },
    ];

    for (const [index, hyperlink] of invalid.entries()) {
      const before = packageSnapshot(pkg);
      const shapes = slide.shapes;
      expect(() => slide.addText('Invalid hyperlink', { hyperlink } as never)).toThrow();
      expect(() => slide.addRichText([{ runs: [{ text: 'Invalid rich hyperlink' }] }], {
        hyperlink,
      } as never)).toThrow();
      expect(() => slide.addPlaceholder('Invalid placeholder hyperlink', {
        name: 'invalid_hyperlink',
        type: 'title',
        hyperlink,
      } as never)).toThrow();
      expect(packageSnapshot(pkg), `invalid hyperlink ${index}`).toEqual(before);
      expect(slide.shapes).toEqual(shapes);
      expect(slide.shapes.at(-1)).toBe(existing);
    }
    expect(getterCalls).toBe(0);

    const before = packageSnapshot(pkg);
    let rolledBack: ShapeModel | undefined;
    expect(() => pkg.transaction(() => {
      rolledBack = slide.addText('Rollback hyperlink', {
        hyperlink: { url: 'https://rollback.example' },
      });
      throw new Error('restore text hyperlink creation');
    })).toThrow('restore text hyperlink creation');
    expect(packageSnapshot(pkg)).toEqual(before);
    expect(slide.shapes.at(-1)).toBe(existing);
    expect(() => rolledBack!.name).toThrow(ModelParseError);
  });

  it('reads and edits direct shape fills through stable live models', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const shape = slide.addShape('rect', {
      fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
        transparency: 25,
      },
    });
    const text = slide.addText('Keep text');

    expect(shape.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF0000' },
      transparency: 25,
    });
    expect(text.fill).toEqual({ kind: 'none' });
    const first = shape.fill;
    const second = shape.fill;
    expect(first).not.toBe(second);
    if (first?.kind === 'solid' && second?.kind === 'solid') {
      expect(first.color).not.toBe(second.color);
    }

    const beforeNoOp = pkg.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...pkg.mutations];
    shape.fill = {
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF0000' },
      transparency: 25,
    };
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeNoOp);
    expect(pkg.mutations).toEqual(noOpJournal);

    shape.fill = { kind: 'none' };
    expect(shape.fill).toEqual({ kind: 'none' });
    shape.fill = undefined;
    expect(shape.fill).toBeUndefined();
    shape.fill = {
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent3' },
      transparency: 40,
    };
    text.fill = {
      kind: 'solid',
      color: { kind: 'srgb', value: '112233' },
    };
    expect(shape.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent3' },
      transparency: 40,
    });
    expect(text.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '112233' },
    });
    expect(slide.shapes[0]).toBe(shape);
    expect(slide.shapes[1]).toBe(text);
    expect(text.text).toBe('Keep text');
  });

  it('replaces and clears unique gradient shape fills and rejects unsafe edits atomically', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const shape = slide.addShape('rect');
    const part = pkg.requirePart(slide.partUri);
    const gradient = new TextDecoder().decode(part.bytes).replace(
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/>',
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '<a:gradFill><a:gsLst/></a:gradFill><a:ln/>',
    );
    pkg.setPart(slide.partUri, gradient, part.contentType);
    expect(shape.fill).toBeUndefined();
    shape.fill = { kind: 'none' };
    expect(shape.fill).toEqual({ kind: 'none' });

    const nonePart = pkg.requirePart(slide.partUri);
    const gradientAgain = new TextDecoder().decode(nonePart.bytes).replace(
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln/>',
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '<a:gradFill><a:gsLst/></a:gradFill><a:ln/>',
    );
    pkg.setPart(slide.partUri, gradientAgain, nonePart.contentType);
    shape.fill = undefined;
    expect(shape.fill).toBeUndefined();
    expect(new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes))
      .not.toContain('<a:gradFill>');

    const ambiguousPart = pkg.requirePart(slide.partUri);
    const ambiguous = new TextDecoder().decode(ambiguousPart.bytes).replace(
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln/>',
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '<a:noFill/><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln/>',
    );
    pkg.setPart(slide.partUri, ambiguous, ambiguousPart.contentType);
    const before = pkg.requirePart(slide.partUri).bytes.slice();
    const parts = pkg.parts.map(({ uri, contentType, bytes }) => ({
      uri,
      contentType,
      bytes: bytes.slice(),
    }));
    const relationships = slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }));
    const journal = [...pkg.mutations];
    expect(shape.fill).toBeUndefined();
    expect(() => {
      shape.fill = { kind: 'none' };
    }).toThrow(ModelParseError);

    let calls = 0;
    const getter = Object.defineProperty({}, 'kind', {
      enumerable: true,
      get() {
        calls += 1;
        return 'none';
      },
    });
    for (const value of [
      null,
      getter,
      { kind: 'solid' },
      { kind: 'solid', color: { kind: 'srgb', value: 'FFF' } },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FFFFFF' },
        transparency: 101,
      },
      { kind: 'none', alpha: 20 },
    ]) {
      expect(() => {
        shape.fill = value as never;
      }).toThrow();
    }
    expect(calls).toBe(0);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(before);
    expect(pkg.parts.map(({ uri, contentType, bytes }) => ({
      uri,
      contentType,
      bytes,
    }))).toEqual(parts);
    expect(slide.relationships.map(({ id, type, target, targetMode }) => ({
      id,
      type,
      target,
      targetMode,
    }))).toEqual(relationships);
    expect(slide.shapes[0]).toBe(shape);
    expect(pkg.mutations).toEqual(journal);
  });

  it('rejects unsafe preset shape trees and id allocation states without mutation', async () => {
    const presentationNamespace =
      'http://schemas.openxmlformats.org/presentationml/2006/main';
    const invalidSources = [
      `<p:sld xmlns:p="${presentationNamespace}"><p:cSld/></p:sld>`,
      `<p:sld xmlns:p="${presentationNamespace}"><p:cSld><p:spTree/>` +
        '<p:spTree/></p:cSld></p:sld>',
      `<p:sld xmlns:p="${presentationNamespace}"><p:cSld><p:spTree>` +
        '<p:extLst/><p:extLst/></p:spTree></p:cSld></p:sld>',
    ];
    for (const id of ['-1', '1.5', '9007199254740992', '4294967295']) {
      invalidSources.push(
        `<p:sld xmlns:p="${presentationNamespace}"><p:cSld><p:spTree>` +
        `<p:nvGrpSpPr><p:cNvPr id="${id}"/></p:nvGrpSpPr>` +
        '</p:spTree></p:cSld></p:sld>',
      );
    }
    invalidSources.push(
      `<p:sld xmlns:p="${presentationNamespace}"><p:cSld><p:spTree>` +
      '<p:nvGrpSpPr><p:cNvPr id="1"/></p:nvGrpSpPr>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="1"/></p:nvSpPr></p:sp>' +
      '</p:spTree></p:cSld></p:sld>',
    );
    invalidSources.push(
      `<p:sld xmlns:p="${presentationNamespace}" xmlns:x="urn:unsafe"><p:cSld><p:spTree>` +
      '<p:nvGrpSpPr><p:cNvPr id="1" x:id="2"/></p:nvGrpSpPr>' +
      '</p:spTree></p:cSld></p:sld>',
    );

    for (const source of invalidSources) {
      const pkg = await OpcPackage.open(await modelFixture());
      const model = new PresentationModel(pkg);
      const slide = model.addSlide();
      const part = pkg.requirePart(slide.partUri);
      pkg.setPart(slide.partUri, source, part.contentType);
      const before = pkg.requirePart(slide.partUri).bytes.slice();
      const journal = [...pkg.mutations];
      expect(() => slide.addShape('ellipse'), source).toThrow(ModelParseError);
      expect(pkg.requirePart(slide.partUri).bytes, source).toEqual(before);
      expect(pkg.mutations, source).toEqual(journal);
    }
  });

  it('reads canonical preset types and rejects ambiguous geometry edits atomically', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const shape = slide.addShape('rect');
    expect(shape.presetType).toBe('rect');

    const part = pkg.requirePart(slide.partUri);
    const ambiguous = new TextDecoder().decode(part.bytes).replace(
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>',
    );
    pkg.setPart(slide.partUri, ambiguous, part.contentType);
    expect(shape.presetType).toBeUndefined();
    const before = pkg.requirePart(slide.partUri).bytes.slice();
    const journal = [...pkg.mutations];
    expect(() => {
      shape.presetType = 'star5';
    }).toThrow(ModelParseError);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(before);
    expect(pkg.mutations).toEqual(journal);
  });

  it('rejects an unsupported presentation content type without guessing from a file name', async () => {
    const pkg = await OpcPackage.open(await modelFixture('application/vnd.example.presentation+xml'));
    let thrown: unknown;
    try {
      new PresentationModel(pkg);
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(UnsupportedPresentationFormatError);
    expect(thrown).toMatchObject({
      contentType: 'application/vnd.example.presentation+xml',
      partUri: '/ppt/presentation.xml',
    });
  });

  it('reads and losslessly edits hidden slide state through the live model', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const identity = model.slides.at(-1);
    expect(identity).toBe(slide);
    expect(slide.hidden).toBe(false);

    const visibleBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const visibleJournal = [...pkg.mutations];
    slide.hidden = false;
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(visibleBytes);
    expect(pkg.mutations).toEqual(visibleJournal);

    const beforeHiddenParts = new Map(
      pkg.parts.map(({ uri, bytes }) => [uri, bytes.slice()]),
    );
    slide.hidden = true;
    expect(slide.hidden).toBe(true);
    expect(model.slides.at(-1)).toBe(identity);
    for (const { uri, bytes } of pkg.parts) {
      if (uri !== slide.partUri) expect(bytes).toEqual(beforeHiddenParts.get(uri));
    }
    let source = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(source).toContain('<p:sld ');
    expect(source).toContain(' show="0">');

    const hiddenBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const hiddenJournal = [...pkg.mutations];
    slide.hidden = true;
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(hiddenBytes);
    expect(pkg.mutations).toEqual(hiddenJournal);

    const tokens: readonly (readonly [string, boolean | undefined])[] = [
      ['0', true],
      ['false', true],
      ['off', true],
      ['1', false],
      ['true', false],
      ['on', false],
      ['maybe', undefined],
    ];
    for (const [token, expected] of tokens) {
      const part = pkg.requirePart(slide.partUri);
      source = new TextDecoder().decode(part.bytes).replace(/ show="[^"]*"/, ` show="${token}"`);
      pkg.setPart(part.uri, source, part.contentType);
      const journal = [...pkg.mutations];
      expect(slide.hidden, token).toBe(expected);
      expect(pkg.mutations).toEqual(journal);
    }

    slide.hidden = true;
    source = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(source).toContain(' show="0">');
    expect(source).not.toContain('show="maybe"');
    slide.hidden = false;
    expect(slide.hidden).toBe(false);
    expect(new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes)).not.toMatch(/\sshow=/);

    const beforeInvalid = pkg.requirePart(slide.partUri).bytes.slice();
    const invalidJournal = [...pkg.mutations];
    for (const invalid of [undefined, null, 0, 1, 'true', {}, [], Symbol('hidden')]) {
      expect(() => {
        (slide as unknown as { hidden: unknown }).hidden = invalid;
      }, String(invalid)).toThrow(TypeError);
    }
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);
    expect(model.slides.at(-1)).toBe(identity);
  });

  it('creates, reads, edits, empties, and clears speaker notes through the live model', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const masterPart = pkg.requirePart('/ppt/notesMasters/notesMaster1.xml');
    pkg.setPart(
      masterPart.uri,
      '<p:notesMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:notesMaster>',
      masterPart.contentType,
    );
    pkg.addRelationship('/ppt/presentation.xml', {
      id: 'rIdNotesMaster',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster',
      target: 'notesMasters/notesMaster1.xml',
    });
    const presentationPart = pkg.requirePart('/ppt/presentation.xml');
    pkg.setPart(
      presentationPart.uri,
      new TextDecoder().decode(presentationPart.bytes).replace(
        '</p:sldIdLst>',
        '</p:sldIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rIdNotesMaster"/></p:notesMasterIdLst>',
      ),
      presentationPart.contentType,
    );
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    expect(slide.notes).toBeUndefined();
    expect(slide.addNotes('First\r\nSecond')).toBe(slide);
    expect(slide.notes).toBe('First\nSecond');

    const notesRelationship = slide.relationships.find(
      ({ type }) => type.endsWith('/notesSlide'),
    );
    const notesUri = notesRelationship?.resolvedTarget;
    expect(notesUri).toBeDefined();
    expect(pkg.requirePart(notesUri!).contentType).toBe(
      'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
    );
    expect(pkg.relationships(notesUri!).find(
      ({ type }) => type.endsWith('/slide'),
    )?.resolvedTarget).toBe(slide.partUri);
    expect(pkg.relationships(notesUri!).find(
      ({ type }) => type.endsWith('/notesMaster'),
    )?.resolvedTarget).toBe('/ppt/notesMasters/notesMaster1.xml');

    const sameParts = pkg.parts.map(({ uri, bytes }) => ({ uri, bytes: bytes.slice() }));
    const sameJournal = [...pkg.mutations];
    slide.notes = 'First\nSecond';
    expect(pkg.parts.map(({ uri, bytes }) => ({ uri, bytes }))).toEqual(sameParts);
    expect(pkg.mutations).toEqual(sameJournal);

    slide.notes = 'Edited';
    expect(slide.notes).toBe('Edited');
    slide.notes = '';
    expect(slide.notes).toBe('');

    const invalidParts = pkg.parts.map(({ uri, bytes }) => ({ uri, bytes: bytes.slice() }));
    const invalidJournal = [...pkg.mutations];
    for (const invalid of [null, 7, true, {}, [], Symbol('notes')]) {
      expect(() => {
        (slide as unknown as { notes: unknown }).notes = invalid;
      }, String(invalid)).toThrow(TypeError);
      expect(() => slide.addNotes(invalid as never), String(invalid)).toThrow(TypeError);
    }
    expect(() => slide.addNotes(undefined as never)).toThrow(TypeError);
    expect(slide.notes).toBe('');
    expect(pkg.parts.map(({ uri, bytes }) => ({ uri, bytes }))).toEqual(invalidParts);
    expect(pkg.mutations).toEqual(invalidJournal);

    slide.notes = 'Reopened';
    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    expect(reopened.slides.at(-1)?.notes).toBe('Reopened');

    slide.notes = undefined;
    expect(slide.notes).toBeUndefined();
    expect(pkg.hasPart(notesUri!)).toBe(false);
    expect(pkg.hasPart('/ppt/notesMasters/notesMaster1.xml')).toBe(true);
    const clearParts = pkg.parts.map(({ uri, bytes }) => ({ uri, bytes: bytes.slice() }));
    const clearJournal = [...pkg.mutations];
    slide.notes = undefined;
    expect(pkg.parts.map(({ uri, bytes }) => ({ uri, bytes }))).toEqual(clearParts);
    expect(pkg.mutations).toEqual(clearJournal);
  });

  it('creates a missing notes master before adding speaker notes', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    pkg.deletePart('/ppt/notesSlides/notesSlide1.xml');
    pkg.deletePart('/ppt/notesMasters/notesMaster1.xml');
    pkg.setPart(
      '/ppt/theme/theme1.xml',
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Notes theme"/>',
      'application/vnd.openxmlformats-officedocument.theme+xml',
    );
    pkg.addRelationship('/ppt/presentation.xml', {
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
      target: 'theme/theme1.xml',
    });

    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    slide.addNotes('Created with master');
    expect(slide.notes).toBe('Created with master');
    const masterRelationships = pkg.relationships('/ppt/presentation.xml').filter(
      ({ type }) => type.endsWith('/notesMaster'),
    );
    expect(masterRelationships).toHaveLength(1);
    const masterUri = masterRelationships[0]!.resolvedTarget!;
    expect(pkg.requirePart(masterUri).contentType).toBe(
      'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
    );
    expect(pkg.relationships(masterUri).find(
      ({ type }) => type.endsWith('/theme'),
    )?.resolvedTarget).toBe('/ppt/theme/theme1.xml');
    expect(new TextDecoder().decode(pkg.requirePart('/ppt/presentation.xml').bytes)).toContain(
      `<p:notesMasterId r:id="${masterRelationships[0]!.id}"/>`,
    );
  });

  it('reads only recognized direct presentation RTL tokens without mutating the package', async () => {
    const cases: readonly [string | undefined, boolean | undefined][] = [
      ['1', true],
      ['true', true],
      ['on', true],
      ['0', false],
      ['false', false],
      ['off', false],
      [undefined, undefined],
      ['', undefined],
      ['yes', undefined],
    ];
    for (const [token, expected] of cases) {
      const pkg = await OpcPackage.open(await modelFixture());
      const part = pkg.requirePart('/ppt/presentation.xml');
      const rtl = token === undefined ? '' : ` rtl="${token}"`;
      pkg.setPart(
        part.uri,
        new TextDecoder().decode(part.bytes).replace('<p:presentation ', `<p:presentation${rtl} `),
        part.contentType,
      );
      const model = new PresentationModel(pkg);
      const journal = [...pkg.mutations];

      expect(model.rtlMode).toBe(expected);
      expect(pkg.mutations).toEqual(journal);
    }

    const descendantPkg = await OpcPackage.open(await modelFixture());
    const descendantPart = descendantPkg.requirePart('/ppt/presentation.xml');
    descendantPkg.setPart(
      descendantPart.uri,
      new TextDecoder().decode(descendantPart.bytes)
        .replace('xmlns:r="r"', 'xmlns:r="r" xmlns:a="a"')
        .replace('</p:presentation>', '<a:lvl1pPr rtl="1"/></p:presentation>'),
      descendantPart.contentType,
    );
    const descendantModel = new PresentationModel(descendantPkg);
    const journal = [...descendantPkg.mutations];
    expect(descendantModel.rtlMode).toBeUndefined();
    expect(descendantPkg.mutations).toEqual(journal);
  });

  it('losslessly replaces, clears, validates, and rolls back presentation RTL', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const part = pkg.requirePart('/ppt/presentation.xml');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<p:presentation ',
        '<p:presentation rtl="yes" saveSubsetFonts="1" custom="KEEP" ',
      ),
      part.contentType,
    );
    const model = new PresentationModel(pkg);
    const slide = model.slides[0];
    const beforeInvalid = pkg.requirePart(part.uri).bytes;
    const journal = [...pkg.mutations];

    for (const invalid of [null, 0, 'true', {}, [], Symbol('rtl')]) {
      expect(() => {
        model.rtlMode = invalid as never;
      }).toThrow(TypeError);
    }
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(journal);
    expect(model.slides[0]).toBe(slide);

    model.rtlMode = true;
    expect(model.rtlMode).toBe(true);
    let updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain(
      '<p:presentation rtl="1" saveSubsetFonts="1" custom="KEEP" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    );
    expect(updated).toContain('<x:unknown xmlns:x="urn:test"/>');

    model.rtlMode = false;
    expect(model.rtlMode).toBe(false);
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain(
      '<p:presentation rtl="0" saveSubsetFonts="1" custom="KEEP" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    );

    const beforeRollback = pkg.requirePart(part.uri).bytes;
    const rollbackJournal = [...pkg.mutations];
    expect(() =>
      pkg.transaction(() => {
        model.rtlMode = true;
        throw new Error('restore presentation RTL');
      }),
    ).toThrow('restore presentation RTL');
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(model.rtlMode).toBe(false);
    expect(model.slides[0]).toBe(slide);

    model.rtlMode = undefined;
    expect(model.rtlMode).toBeUndefined();
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain(
      '<p:presentation saveSubsetFonts="1" custom="KEEP" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    );
    expect(updated).not.toMatch(/<p:presentation[^>]*\srtl=/);
    expect(updated).toContain('<x:unknown xmlns:x="urn:test"/>');
  });

  it('reads, edits, clears, creates, and reopens presentation title metadata', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const coreXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:coreProperties xmlns:c="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:d="http://purl.org/dc/elements/1.1/" xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">' +
      '<d:title>Original &amp; title</d:title><cp:revision>7</cp:revision>' +
      '<x:opaque xmlns:x="urn:test">KEEP</x:opaque></c:coreProperties>';
    pkg.transaction(() => {
      pkg.setPart('/metadata/properties.xml', coreXml, CORE_PROPERTIES_CONTENT_TYPE);
      pkg.addRelationship('/', {
        type: CORE_PROPERTIES_RELATIONSHIP,
        target: 'metadata/properties.xml',
      });
    });
    const model = new PresentationModel(pkg);
    const slide = model.slides[0];
    const readJournal = [...pkg.mutations];
    expect(model.title).toBe('Original & title');
    expect(pkg.mutations).toEqual(readJournal);

    const beforeSame = pkg.requirePart('/metadata/properties.xml').bytes;
    model.title = 'Original & title';
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeSame);
    expect(pkg.mutations).toEqual(readJournal);
    expect(model.slides[0]).toBe(slide);

    const beforeInvalid = pkg.requirePart('/metadata/properties.xml').bytes;
    const invalidJournal = [...pkg.mutations];
    for (const value of [null, false, 0, {}, [], Symbol('title'), 'bad\u0001title']) {
      expect(() => {
        model.title = value as never;
      }).toThrow(TypeError);
    }
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);
    expect(model.slides[0]).toBe(slide);

    const otherParts = new Map(
      pkg.parts
        .filter(({ uri }) => uri !== '/metadata/properties.xml')
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    model.title = 'Edited & <safe>';
    expect(model.title).toBe('Edited & <safe>');
    let updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).toContain('<d:title>Edited &amp; &lt;safe&gt;</d:title>');
    expect(updated).toContain('<cp:revision>7</cp:revision>');
    expect(updated).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    for (const [uri, bytes] of otherParts) {
      expect(pkg.requirePart(uri).bytes).toEqual(bytes);
    }
    expect(model.slides[0]).toBe(slide);

    model.title = '';
    expect(model.title).toBe('');
    updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).toContain('<d:title></d:title>');

    const beforeRollback = pkg.requirePart('/metadata/properties.xml').bytes;
    const rollbackJournal = [...pkg.mutations];
    expect(() => pkg.transaction(() => {
      model.title = 'Temporary';
      expect(model.title).toBe('Temporary');
      throw new Error('restore presentation title');
    })).toThrow('restore presentation title');
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(model.title).toBe('');

    model.title = undefined;
    expect(model.title).toBeUndefined();
    updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).not.toContain('<d:title');
    expect(updated).toContain('<cp:revision>7</cp:revision>');
    expect(updated).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    expect(pkg.hasPart('/metadata/properties.xml')).toBe(true);
    expect(pkg.relationships('/').filter(
      ({ type }) => type === CORE_PROPERTIES_RELATIONSHIP,
    )).toHaveLength(1);

    model.title = 'Reopened title';
    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    expect(reopened.title).toBe('Reopened title');
    expect(reopened.slides.map(({ partUri }) => partUri)).toEqual(
      model.slides.map(({ partUri }) => partUri),
    );

    const missingPkg = await OpcPackage.open(await modelFixture());
    const missingModel = new PresentationModel(missingPkg);
    const missingSlide = missingModel.slides[0];
    const beforeAbsentClear = missingPkg.parts.map(({ uri, bytes }) => [uri, bytes] as const);
    const absentJournal = [...missingPkg.mutations];
    expect(missingModel.title).toBeUndefined();
    missingModel.title = undefined;
    expect(missingPkg.mutations).toEqual(absentJournal);
    for (const [uri, bytes] of beforeAbsentClear) {
      expect(missingPkg.requirePart(uri).bytes).toEqual(bytes);
    }

    missingModel.title = 'Created metadata';
    expect(missingModel.title).toBe('Created metadata');
    expect(missingModel.slides[0]).toBe(missingSlide);
    const createdRelationship = missingPkg.relationships('/').find(
      ({ type }) => type === CORE_PROPERTIES_RELATIONSHIP,
    );
    expect(createdRelationship?.resolvedTarget).toBe('/docProps/core.xml');
    expect(missingPkg.requirePart('/docProps/core.xml').contentType)
      .toBe(CORE_PROPERTIES_CONTENT_TYPE);
    const reopenedCreated = new PresentationModel(await OpcPackage.open(await missingPkg.write()));
    expect(reopenedCreated.title).toBe('Created metadata');
  });

  it('reads, edits, clears, creates, and reopens presentation author metadata', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const coreXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:coreProperties xmlns:c="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:d="http://purl.org/dc/elements/1.1/" xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">' +
      '<d:creator>Original &amp; author</d:creator><d:title>Quarterly</d:title>' +
      '<cp:lastModifiedBy>Editor</cp:lastModifiedBy><cp:revision>7</cp:revision>' +
      '<x:opaque xmlns:x="urn:test">KEEP</x:opaque></c:coreProperties>';
    pkg.transaction(() => {
      pkg.setPart('/metadata/properties.xml', coreXml, CORE_PROPERTIES_CONTENT_TYPE);
      pkg.addRelationship('/', {
        type: CORE_PROPERTIES_RELATIONSHIP,
        target: 'metadata/properties.xml',
      });
    });
    const model = new PresentationModel(pkg);
    const slide = model.slides[0];
    const readJournal = [...pkg.mutations];
    expect(model.author).toBe('Original & author');
    expect(pkg.mutations).toEqual(readJournal);

    const beforeSame = pkg.requirePart('/metadata/properties.xml').bytes;
    model.author = 'Original & author';
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeSame);
    expect(pkg.mutations).toEqual(readJournal);
    expect(model.slides[0]).toBe(slide);

    const beforeInvalid = pkg.requirePart('/metadata/properties.xml').bytes;
    const invalidJournal = [...pkg.mutations];
    for (const value of [null, false, 0, {}, [], Symbol('author'), 'bad\u0001author']) {
      expect(() => {
        model.author = value as never;
      }).toThrow(TypeError);
    }
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);
    expect(model.slides[0]).toBe(slide);

    const otherParts = new Map(
      pkg.parts
        .filter(({ uri }) => uri !== '/metadata/properties.xml')
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    model.author = 'Edited & <safe>';
    expect(model.author).toBe('Edited & <safe>');
    let updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).toContain('<d:creator>Edited &amp; &lt;safe&gt;</d:creator>');
    expect(updated).toContain('<d:title>Quarterly</d:title>');
    expect(updated).toContain('<cp:lastModifiedBy>Editor</cp:lastModifiedBy>');
    expect(updated).toContain('<cp:revision>7</cp:revision>');
    expect(updated).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    for (const [uri, bytes] of otherParts) {
      expect(pkg.requirePart(uri).bytes).toEqual(bytes);
    }
    expect(model.slides[0]).toBe(slide);

    model.author = '';
    expect(model.author).toBe('');
    updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).toContain('<d:creator></d:creator>');
    expect(updated).toContain('<cp:lastModifiedBy>Editor</cp:lastModifiedBy>');

    const beforeRollback = pkg.requirePart('/metadata/properties.xml').bytes;
    const rollbackJournal = [...pkg.mutations];
    expect(() => pkg.transaction(() => {
      model.author = 'Temporary';
      expect(model.author).toBe('Temporary');
      throw new Error('restore presentation author');
    })).toThrow('restore presentation author');
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(model.author).toBe('');

    model.author = undefined;
    expect(model.author).toBeUndefined();
    updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).not.toContain('<d:creator');
    expect(updated).toContain('<d:title>Quarterly</d:title>');
    expect(updated).toContain('<cp:lastModifiedBy>Editor</cp:lastModifiedBy>');
    expect(updated).toContain('<cp:revision>7</cp:revision>');
    expect(updated).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    expect(pkg.hasPart('/metadata/properties.xml')).toBe(true);
    expect(pkg.relationships('/').filter(
      ({ type }) => type === CORE_PROPERTIES_RELATIONSHIP,
    )).toHaveLength(1);

    model.author = 'Reopened author';
    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    expect(reopened.author).toBe('Reopened author');
    expect(reopened.slides.map(({ partUri }) => partUri)).toEqual(
      model.slides.map(({ partUri }) => partUri),
    );
    expect(new TextDecoder().decode(
      reopened.opcPackage.requirePart('/metadata/properties.xml').bytes,
    )).toContain('<cp:lastModifiedBy>Editor</cp:lastModifiedBy>');

    const missingPkg = await OpcPackage.open(await modelFixture());
    const missingModel = new PresentationModel(missingPkg);
    const missingSlide = missingModel.slides[0];
    const beforeAbsentClear = missingPkg.parts.map(({ uri, bytes }) => [uri, bytes] as const);
    const absentJournal = [...missingPkg.mutations];
    expect(missingModel.author).toBeUndefined();
    missingModel.author = undefined;
    expect(missingPkg.mutations).toEqual(absentJournal);
    for (const [uri, bytes] of beforeAbsentClear) {
      expect(missingPkg.requirePart(uri).bytes).toEqual(bytes);
    }

    missingModel.author = 'Created metadata';
    expect(missingModel.author).toBe('Created metadata');
    expect(missingModel.slides[0]).toBe(missingSlide);
    const createdRelationship = missingPkg.relationships('/').find(
      ({ type }) => type === CORE_PROPERTIES_RELATIONSHIP,
    );
    expect(createdRelationship?.resolvedTarget).toBe('/docProps/core.xml');
    expect(missingPkg.requirePart('/docProps/core.xml').contentType)
      .toBe(CORE_PROPERTIES_CONTENT_TYPE);
    const createdXml = new TextDecoder().decode(
      missingPkg.requirePart('/docProps/core.xml').bytes,
    );
    expect(createdXml).toContain('<dc:creator>Created metadata</dc:creator>');
    expect(createdXml).not.toContain('lastModifiedBy');
    const reopenedCreated = new PresentationModel(await OpcPackage.open(await missingPkg.write()));
    expect(reopenedCreated.author).toBe('Created metadata');
  });

  it('reads, edits, clears, creates, and reopens presentation last modified by metadata', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const coreXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:coreProperties xmlns:c="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:d="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
      'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<d:creator>Alice</d:creator><c:lastModifiedBy>Original &amp; editor</c:lastModifiedBy>' +
      '<d:title>Quarterly</d:title><d:subject>Forecast</d:subject><c:revision>7</c:revision>' +
      '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>' +
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-30T01:00:00Z</dcterms:modified>' +
      '<x:opaque xmlns:x="urn:test">KEEP</x:opaque></c:coreProperties>';
    pkg.transaction(() => {
      pkg.setPart('/metadata/properties.xml', coreXml, CORE_PROPERTIES_CONTENT_TYPE);
      pkg.addRelationship('/', {
        type: CORE_PROPERTIES_RELATIONSHIP,
        target: 'metadata/properties.xml',
      });
    });
    const model = new PresentationModel(pkg);
    const slide = model.slides[0];
    const readJournal = [...pkg.mutations];
    expect(model.lastModifiedBy).toBe('Original & editor');
    expect(pkg.mutations).toEqual(readJournal);

    const beforeSame = pkg.requirePart('/metadata/properties.xml').bytes;
    model.lastModifiedBy = 'Original & editor';
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeSame);
    expect(pkg.mutations).toEqual(readJournal);
    expect(model.slides[0]).toBe(slide);

    const beforeInvalid = pkg.requirePart('/metadata/properties.xml').bytes;
    const invalidJournal = [...pkg.mutations];
    for (const value of [
      null, true, false, 0, 1n, {}, [], Symbol('lastModifiedBy'), 'bad\u0001editor',
    ]) {
      expect(() => {
        model.lastModifiedBy = value as never;
      }).toThrow(TypeError);
    }
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);
    expect(model.slides[0]).toBe(slide);

    const otherParts = new Map(
      pkg.parts
        .filter(({ uri }) => uri !== '/metadata/properties.xml')
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    model.lastModifiedBy = 'Edited & <safe>';
    expect(model.lastModifiedBy).toBe('Edited & <safe>');
    let updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).toContain(
      '<c:lastModifiedBy>Edited &amp; &lt;safe&gt;</c:lastModifiedBy>',
    );
    expect(updated).toContain('<d:creator>Alice</d:creator>');
    expect(updated).toContain('<d:title>Quarterly</d:title>');
    expect(updated).toContain('<d:subject>Forecast</d:subject>');
    expect(updated).toContain('<c:revision>7</c:revision>');
    expect(updated).toContain('2026-07-30T00:00:00Z');
    expect(updated).toContain('2026-07-30T01:00:00Z');
    expect(updated).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    for (const [uri, bytes] of otherParts) {
      expect(pkg.requirePart(uri).bytes).toEqual(bytes);
    }
    expect(model.slides[0]).toBe(slide);

    model.author = 'Changed author';
    expect(model.author).toBe('Changed author');
    expect(model.lastModifiedBy).toBe('Edited & <safe>');
    model.lastModifiedBy = '';
    expect(model.lastModifiedBy).toBe('');
    expect(model.author).toBe('Changed author');
    updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).toContain('<c:lastModifiedBy></c:lastModifiedBy>');
    expect(updated).toContain('<d:creator>Changed author</d:creator>');

    const beforeRollback = pkg.requirePart('/metadata/properties.xml').bytes;
    const rollbackJournal = [...pkg.mutations];
    expect(() => pkg.transaction(() => {
      model.lastModifiedBy = 'Temporary';
      expect(model.lastModifiedBy).toBe('Temporary');
      throw new Error('restore presentation lastModifiedBy');
    })).toThrow('restore presentation lastModifiedBy');
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(model.lastModifiedBy).toBe('');

    model.lastModifiedBy = undefined;
    expect(model.lastModifiedBy).toBeUndefined();
    updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).not.toContain('<c:lastModifiedBy');
    expect(updated).toContain('<d:creator>Changed author</d:creator>');
    expect(updated).toContain('<d:title>Quarterly</d:title>');
    expect(updated).toContain('<c:revision>7</c:revision>');
    expect(updated).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    expect(pkg.hasPart('/metadata/properties.xml')).toBe(true);

    model.lastModifiedBy = 'Reopened editor';
    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    expect(reopened.lastModifiedBy).toBe('Reopened editor');
    expect(reopened.author).toBe('Changed author');
    expect(reopened.slides.map(({ partUri }) => partUri)).toEqual(
      model.slides.map(({ partUri }) => partUri),
    );

    const missingPkg = await OpcPackage.open(await modelFixture());
    const missingModel = new PresentationModel(missingPkg);
    const missingSlide = missingModel.slides[0];
    const beforeAbsentClear = missingPkg.parts.map(({ uri, bytes }) => [uri, bytes] as const);
    const absentJournal = [...missingPkg.mutations];
    expect(missingModel.lastModifiedBy).toBeUndefined();
    missingModel.lastModifiedBy = undefined;
    expect(missingPkg.mutations).toEqual(absentJournal);
    for (const [uri, bytes] of beforeAbsentClear) {
      expect(missingPkg.requirePart(uri).bytes).toEqual(bytes);
    }

    missingModel.lastModifiedBy = 'Created metadata';
    expect(missingModel.lastModifiedBy).toBe('Created metadata');
    expect(missingModel.slides[0]).toBe(missingSlide);
    const createdRelationship = missingPkg.relationships('/').find(
      ({ type }) => type === CORE_PROPERTIES_RELATIONSHIP,
    );
    expect(createdRelationship?.resolvedTarget).toBe('/docProps/core.xml');
    expect(missingPkg.requirePart('/docProps/core.xml').contentType)
      .toBe(CORE_PROPERTIES_CONTENT_TYPE);
    const createdXml = new TextDecoder().decode(
      missingPkg.requirePart('/docProps/core.xml').bytes,
    );
    expect(createdXml).toContain('<cp:lastModifiedBy>Created metadata</cp:lastModifiedBy>');
    expect(createdXml.match(/xmlns:cp=/g)).toHaveLength(1);
    expect(createdXml).not.toContain('creator');
    const reopenedCreated = new PresentationModel(await OpcPackage.open(await missingPkg.write()));
    expect(reopenedCreated.lastModifiedBy).toBe('Created metadata');
  });

  it('reads, edits, clears, repairs, creates, and reopens presentation created-at metadata', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const coreXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:coreProperties xmlns:c="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:d="http://purl.org/dc/elements/1.1/" xmlns:t="http://purl.org/dc/terms/" ' +
      'xmlns:i="http://www.w3.org/2001/XMLSchema-instance">' +
      '<d:creator>Alice</d:creator><c:lastModifiedBy>Editor</c:lastModifiedBy>' +
      '<d:title>Quarterly</d:title><d:subject>Forecast</d:subject><c:revision>7</c:revision>' +
      '<t:created i:type="t:W3CDTF">2024-02-29T12:34:56.123456+05:30</t:created>' +
      '<t:modified i:type="t:W3CDTF">2026-07-30T01:00:00Z</t:modified>' +
      '<x:opaque xmlns:x="urn:test">KEEP</x:opaque></c:coreProperties>';
    pkg.transaction(() => {
      pkg.setPart('/metadata/properties.xml', coreXml, CORE_PROPERTIES_CONTENT_TYPE);
      pkg.addRelationship('/', {
        type: CORE_PROPERTIES_RELATIONSHIP,
        target: 'metadata/properties.xml',
      });
    });
    const model = new PresentationModel(pkg);
    const slide = model.slides[0];
    const readJournal = [...pkg.mutations];
    expect(model.createdAt).toBe('2024-02-29T12:34:56.123456+05:30');
    expect(pkg.mutations).toEqual(readJournal);

    const beforeSame = pkg.requirePart('/metadata/properties.xml').bytes;
    model.createdAt = '2024-02-29T12:34:56.123456+05:30';
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeSame);
    expect(pkg.mutations).toEqual(readJournal);
    expect(model.slides[0]).toBe(slide);

    const beforeInvalid = pkg.requirePart('/metadata/properties.xml').bytes;
    const invalidJournal = [...pkg.mutations];
    for (const value of [
      '',
      '1900-02-29T00:00:00Z',
      '2026-07-30T00:00:00',
      '2026-07-30T00:00:00+14:01',
      null,
      false,
      0,
      new Date(),
      {},
      [],
      Symbol('createdAt'),
    ]) {
      expect(() => {
        model.createdAt = value as never;
      }).toThrow(TypeError);
    }
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);
    expect(model.slides[0]).toBe(slide);

    const otherParts = new Map(
      pkg.parts
        .filter(({ uri }) => uri !== '/metadata/properties.xml')
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    model.createdAt = '2026-07-30T00:00:00Z';
    expect(model.createdAt).toBe('2026-07-30T00:00:00Z');
    let updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).toContain(
      '<t:created i:type="t:W3CDTF">2026-07-30T00:00:00Z</t:created>',
    );
    expect(updated).toContain(
      '<t:modified i:type="t:W3CDTF">2026-07-30T01:00:00Z</t:modified>',
    );
    expect(updated).toContain('<d:creator>Alice</d:creator>');
    expect(updated).toContain('<c:lastModifiedBy>Editor</c:lastModifiedBy>');
    expect(updated).toContain('<c:revision>7</c:revision>');
    expect(updated).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    for (const [uri, bytes] of otherParts) {
      expect(pkg.requirePart(uri).bytes).toEqual(bytes);
    }
    expect(model.slides[0]).toBe(slide);

    model.lastModifiedBy = 'Changed editor';
    model.revision = '8';
    expect(model.createdAt).toBe('2026-07-30T00:00:00Z');
    expect(model.lastModifiedBy).toBe('Changed editor');
    expect(model.revision).toBe('8');

    const beforeRollback = pkg.requirePart('/metadata/properties.xml').bytes;
    const rollbackJournal = [...pkg.mutations];
    expect(() => pkg.transaction(() => {
      model.createdAt = '2025-01-01T00:00:00Z';
      expect(model.createdAt).toBe('2025-01-01T00:00:00Z');
      throw new Error('restore presentation createdAt');
    })).toThrow('restore presentation createdAt');
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(model.createdAt).toBe('2026-07-30T00:00:00Z');

    model.createdAt = undefined;
    expect(model.createdAt).toBeUndefined();
    updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).not.toContain('<t:created');
    expect(updated).toContain(
      '<t:modified i:type="t:W3CDTF">2026-07-30T01:00:00Z</t:modified>',
    );
    expect(updated).toContain('<c:lastModifiedBy>Changed editor</c:lastModifiedBy>');
    expect(updated).toContain('<c:revision>8</c:revision>');
    expect(pkg.hasPart('/metadata/properties.xml')).toBe(true);

    model.createdAt = '2024-02-29T12:34:56Z';
    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    expect(reopened.createdAt).toBe('2024-02-29T12:34:56Z');
    expect(reopened.lastModifiedBy).toBe('Changed editor');
    expect(reopened.revision).toBe('8');
    expect(reopened.slides.map(({ partUri }) => partUri)).toEqual(
      model.slides.map(({ partUri }) => partUri),
    );

    const repairPkg = await OpcPackage.open(await modelFixture());
    repairPkg.transaction(() => {
      repairPkg.setPart(
        '/metadata/repair-core.xml',
        '<?xml version="1.0"?><c:coreProperties ' +
          'xmlns:c="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
          'xmlns:t="http://purl.org/dc/terms/" ' +
          'xmlns:i="http://www.w3.org/2001/XMLSchema-instance">' +
          '<t:created i:type="t:Other">invalid</t:created>' +
          '<t:modified i:type="t:W3CDTF">2026-07-30T01:00:00Z</t:modified>' +
          '</c:coreProperties>',
        CORE_PROPERTIES_CONTENT_TYPE,
      );
      repairPkg.addRelationship('/', {
        type: CORE_PROPERTIES_RELATIONSHIP,
        target: 'metadata/repair-core.xml',
      });
    });
    const repairModel = new PresentationModel(repairPkg);
    expect(repairModel.createdAt).toBeUndefined();
    repairModel.createdAt = '2026-07-30T00:00:00Z';
    expect(repairModel.createdAt).toBe('2026-07-30T00:00:00Z');
    const repairedXml = new TextDecoder().decode(
      repairPkg.requirePart('/metadata/repair-core.xml').bytes,
    );
    expect(repairedXml).toContain(
      '<t:created i:type="t:W3CDTF">2026-07-30T00:00:00Z</t:created>',
    );
    expect(repairedXml).toContain(
      '<t:modified i:type="t:W3CDTF">2026-07-30T01:00:00Z</t:modified>',
    );

    const missingPkg = await OpcPackage.open(await modelFixture());
    const missingModel = new PresentationModel(missingPkg);
    const missingSlide = missingModel.slides[0];
    const beforeAbsentClear = missingPkg.parts.map(({ uri, bytes }) => [uri, bytes] as const);
    const absentJournal = [...missingPkg.mutations];
    expect(missingModel.createdAt).toBeUndefined();
    missingModel.createdAt = undefined;
    expect(missingPkg.mutations).toEqual(absentJournal);
    for (const [uri, bytes] of beforeAbsentClear) {
      expect(missingPkg.requirePart(uri).bytes).toEqual(bytes);
    }

    missingModel.createdAt = '2024-02-29T12:34:56.123+05:30';
    expect(missingModel.createdAt).toBe('2024-02-29T12:34:56.123+05:30');
    expect(missingModel.slides[0]).toBe(missingSlide);
    const createdRelationship = missingPkg.relationships('/').find(
      ({ type }) => type === CORE_PROPERTIES_RELATIONSHIP,
    );
    expect(createdRelationship?.resolvedTarget).toBe('/docProps/core.xml');
    expect(missingPkg.requirePart('/docProps/core.xml').contentType)
      .toBe(CORE_PROPERTIES_CONTENT_TYPE);
    const createdXml = new TextDecoder().decode(
      missingPkg.requirePart('/docProps/core.xml').bytes,
    );
    expect(createdXml).toContain(
      '<dcterms:created xsi:type="dcterms:W3CDTF">2024-02-29T12:34:56.123+05:30</dcterms:created>',
    );
    expect(createdXml.match(/xmlns:dcterms=/g)).toHaveLength(1);
    expect(createdXml.match(/xmlns:xsi=/g)).toHaveLength(1);
    expect(createdXml).not.toContain('modified');
    const reopenedCreated = new PresentationModel(await OpcPackage.open(await missingPkg.write()));
    expect(reopenedCreated.createdAt).toBe('2024-02-29T12:34:56.123+05:30');
  });

  it('reads, edits, clears, repairs, creates, and reopens presentation modified-at metadata', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const coreXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:coreProperties xmlns:c="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:d="http://purl.org/dc/elements/1.1/" xmlns:t="http://purl.org/dc/terms/" ' +
      'xmlns:i="http://www.w3.org/2001/XMLSchema-instance">' +
      '<d:creator>Alice</d:creator><c:lastModifiedBy>Editor</c:lastModifiedBy>' +
      '<d:title>Quarterly</d:title><d:subject>Forecast</d:subject><c:revision>7</c:revision>' +
      '<t:created i:type="t:W3CDTF">2024-02-29T12:34:56.123456+05:30</t:created>' +
      '<t:modified i:type="t:W3CDTF">2026-07-30T01:00:00Z</t:modified>' +
      '<x:opaque xmlns:x="urn:test">KEEP</x:opaque></c:coreProperties>';
    pkg.transaction(() => {
      pkg.setPart('/metadata/properties.xml', coreXml, CORE_PROPERTIES_CONTENT_TYPE);
      pkg.addRelationship('/', {
        type: CORE_PROPERTIES_RELATIONSHIP,
        target: 'metadata/properties.xml',
      });
    });
    const model = new PresentationModel(pkg);
    const slide = model.slides[0];
    const readJournal = [...pkg.mutations];
    expect(model.createdAt).toBe('2024-02-29T12:34:56.123456+05:30');
    expect(model.modifiedAt).toBe('2026-07-30T01:00:00Z');
    expect(pkg.mutations).toEqual(readJournal);

    const beforeSame = pkg.requirePart('/metadata/properties.xml').bytes;
    model.modifiedAt = '2026-07-30T01:00:00Z';
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeSame);
    expect(pkg.mutations).toEqual(readJournal);
    expect(model.slides[0]).toBe(slide);

    const beforeInvalid = pkg.requirePart('/metadata/properties.xml').bytes;
    const invalidJournal = [...pkg.mutations];
    for (const value of [
      '',
      '1900-02-29T00:00:00Z',
      '2026-07-30T00:00:00',
      '2026-07-30T00:00:00+14:01',
      null,
      false,
      0,
      new Date(),
      {},
      [],
      Symbol('modifiedAt'),
    ]) {
      expect(() => {
        model.modifiedAt = value as never;
      }).toThrow(TypeError);
    }
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);
    expect(model.slides[0]).toBe(slide);

    const otherParts = new Map(
      pkg.parts
        .filter(({ uri }) => uri !== '/metadata/properties.xml')
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    model.modifiedAt = '2026-07-30T02:03:04.5+08:00';
    expect(model.modifiedAt).toBe('2026-07-30T02:03:04.5+08:00');
    expect(model.createdAt).toBe('2024-02-29T12:34:56.123456+05:30');
    let updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).toContain(
      '<t:modified i:type="t:W3CDTF">2026-07-30T02:03:04.5+08:00</t:modified>',
    );
    expect(updated).toContain(
      '<t:created i:type="t:W3CDTF">2024-02-29T12:34:56.123456+05:30</t:created>',
    );
    expect(updated).toContain('<d:creator>Alice</d:creator>');
    expect(updated).toContain('<c:lastModifiedBy>Editor</c:lastModifiedBy>');
    expect(updated).toContain('<c:revision>7</c:revision>');
    expect(updated).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    for (const [uri, bytes] of otherParts) {
      expect(pkg.requirePart(uri).bytes).toEqual(bytes);
    }
    expect(model.slides[0]).toBe(slide);

    model.createdAt = '2024-02-29T12:34:56Z';
    model.lastModifiedBy = 'Changed editor';
    model.revision = '8';
    expect(model.modifiedAt).toBe('2026-07-30T02:03:04.5+08:00');
    expect(model.createdAt).toBe('2024-02-29T12:34:56Z');
    expect(model.lastModifiedBy).toBe('Changed editor');
    expect(model.revision).toBe('8');

    const beforeRollback = pkg.requirePart('/metadata/properties.xml').bytes;
    const rollbackJournal = [...pkg.mutations];
    expect(() => pkg.transaction(() => {
      model.modifiedAt = '2025-01-01T00:00:00Z';
      expect(model.modifiedAt).toBe('2025-01-01T00:00:00Z');
      throw new Error('restore presentation modifiedAt');
    })).toThrow('restore presentation modifiedAt');
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(model.modifiedAt).toBe('2026-07-30T02:03:04.5+08:00');

    model.modifiedAt = undefined;
    expect(model.modifiedAt).toBeUndefined();
    updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).not.toContain('<t:modified');
    expect(updated).toContain(
      '<t:created i:type="t:W3CDTF">2024-02-29T12:34:56Z</t:created>',
    );
    expect(updated).toContain('<c:lastModifiedBy>Changed editor</c:lastModifiedBy>');
    expect(updated).toContain('<c:revision>8</c:revision>');

    model.modifiedAt = '2024-03-01T01:02:03Z';
    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    expect(reopened.modifiedAt).toBe('2024-03-01T01:02:03Z');
    expect(reopened.createdAt).toBe('2024-02-29T12:34:56Z');
    expect(reopened.lastModifiedBy).toBe('Changed editor');
    expect(reopened.revision).toBe('8');
    expect(reopened.slides.map(({ partUri }) => partUri)).toEqual(
      model.slides.map(({ partUri }) => partUri),
    );

    const repairPkg = await OpcPackage.open(await modelFixture());
    repairPkg.transaction(() => {
      repairPkg.setPart(
        '/metadata/repair-core.xml',
        '<?xml version="1.0"?><c:coreProperties ' +
          'xmlns:c="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
          'xmlns:t="http://purl.org/dc/terms/" ' +
          'xmlns:i="http://www.w3.org/2001/XMLSchema-instance">' +
          '<t:created i:type="t:W3CDTF">2026-07-29T00:00:00Z</t:created>' +
          '<t:modified i:type="t:Other">invalid</t:modified>' +
          '</c:coreProperties>',
        CORE_PROPERTIES_CONTENT_TYPE,
      );
      repairPkg.addRelationship('/', {
        type: CORE_PROPERTIES_RELATIONSHIP,
        target: 'metadata/repair-core.xml',
      });
    });
    const repairModel = new PresentationModel(repairPkg);
    expect(repairModel.modifiedAt).toBeUndefined();
    repairModel.modifiedAt = '2026-07-30T01:02:03Z';
    expect(repairModel.modifiedAt).toBe('2026-07-30T01:02:03Z');
    expect(repairModel.createdAt).toBe('2026-07-29T00:00:00Z');
    const repairedXml = new TextDecoder().decode(
      repairPkg.requirePart('/metadata/repair-core.xml').bytes,
    );
    expect(repairedXml).toContain(
      '<t:modified i:type="t:W3CDTF">2026-07-30T01:02:03Z</t:modified>',
    );
    expect(repairedXml).toContain(
      '<t:created i:type="t:W3CDTF">2026-07-29T00:00:00Z</t:created>',
    );

    const missingPkg = await OpcPackage.open(await modelFixture());
    const missingModel = new PresentationModel(missingPkg);
    const missingSlide = missingModel.slides[0];
    const beforeAbsentClear = missingPkg.parts.map(({ uri, bytes }) => [uri, bytes] as const);
    const absentJournal = [...missingPkg.mutations];
    expect(missingModel.modifiedAt).toBeUndefined();
    missingModel.modifiedAt = undefined;
    expect(missingPkg.mutations).toEqual(absentJournal);
    for (const [uri, bytes] of beforeAbsentClear) {
      expect(missingPkg.requirePart(uri).bytes).toEqual(bytes);
    }

    missingModel.modifiedAt = '2024-03-01T01:02:03.456+08:00';
    expect(missingModel.modifiedAt).toBe('2024-03-01T01:02:03.456+08:00');
    expect(missingModel.slides[0]).toBe(missingSlide);
    const createdRelationship = missingPkg.relationships('/').find(
      ({ type }) => type === CORE_PROPERTIES_RELATIONSHIP,
    );
    expect(createdRelationship?.resolvedTarget).toBe('/docProps/core.xml');
    expect(missingPkg.requirePart('/docProps/core.xml').contentType)
      .toBe(CORE_PROPERTIES_CONTENT_TYPE);
    const createdXml = new TextDecoder().decode(
      missingPkg.requirePart('/docProps/core.xml').bytes,
    );
    expect(createdXml).toContain(
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2024-03-01T01:02:03.456+08:00</dcterms:modified>',
    );
    expect(createdXml.match(/xmlns:dcterms=/g)).toHaveLength(1);
    expect(createdXml.match(/xmlns:xsi=/g)).toHaveLength(1);
    expect(createdXml).not.toContain('created');
    const reopenedCreated = new PresentationModel(await OpcPackage.open(await missingPkg.write()));
    expect(reopenedCreated.modifiedAt).toBe('2024-03-01T01:02:03.456+08:00');
  });

  it('reads, edits, clears, creates, and reopens presentation subject metadata', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const coreXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:coreProperties xmlns:c="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:d="http://purl.org/dc/elements/1.1/" xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">' +
      '<d:subject>Original &amp; subject</d:subject><d:title>Quarterly</d:title>' +
      '<d:creator>Alice</d:creator><cp:lastModifiedBy>Editor</cp:lastModifiedBy>' +
      '<cp:revision>7</cp:revision><x:opaque xmlns:x="urn:test">KEEP</x:opaque>' +
      '</c:coreProperties>';
    pkg.transaction(() => {
      pkg.setPart('/metadata/properties.xml', coreXml, CORE_PROPERTIES_CONTENT_TYPE);
      pkg.addRelationship('/', {
        type: CORE_PROPERTIES_RELATIONSHIP,
        target: 'metadata/properties.xml',
      });
    });
    const model = new PresentationModel(pkg);
    const slide = model.slides[0];
    const readJournal = [...pkg.mutations];
    expect(model.subject).toBe('Original & subject');
    expect(pkg.mutations).toEqual(readJournal);

    const beforeSame = pkg.requirePart('/metadata/properties.xml').bytes;
    model.subject = 'Original & subject';
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeSame);
    expect(pkg.mutations).toEqual(readJournal);
    expect(model.slides[0]).toBe(slide);

    const beforeInvalid = pkg.requirePart('/metadata/properties.xml').bytes;
    const invalidJournal = [...pkg.mutations];
    for (const value of [null, false, 0, {}, [], Symbol('subject'), 'bad\u0001subject']) {
      expect(() => {
        model.subject = value as never;
      }).toThrow(TypeError);
    }
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);
    expect(model.slides[0]).toBe(slide);

    const otherParts = new Map(
      pkg.parts
        .filter(({ uri }) => uri !== '/metadata/properties.xml')
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    model.subject = 'Edited & <safe>';
    expect(model.subject).toBe('Edited & <safe>');
    let updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).toContain('<d:subject>Edited &amp; &lt;safe&gt;</d:subject>');
    expect(updated).toContain('<d:title>Quarterly</d:title>');
    expect(updated).toContain('<d:creator>Alice</d:creator>');
    expect(updated).toContain('<cp:lastModifiedBy>Editor</cp:lastModifiedBy>');
    expect(updated).toContain('<cp:revision>7</cp:revision>');
    expect(updated).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    for (const [uri, bytes] of otherParts) {
      expect(pkg.requirePart(uri).bytes).toEqual(bytes);
    }
    expect(model.slides[0]).toBe(slide);

    model.subject = '';
    expect(model.subject).toBe('');
    updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).toContain('<d:subject></d:subject>');
    expect(updated).toContain('<d:title>Quarterly</d:title>');
    expect(updated).toContain('<d:creator>Alice</d:creator>');

    const beforeRollback = pkg.requirePart('/metadata/properties.xml').bytes;
    const rollbackJournal = [...pkg.mutations];
    expect(() => pkg.transaction(() => {
      model.subject = 'Temporary';
      expect(model.subject).toBe('Temporary');
      throw new Error('restore presentation subject');
    })).toThrow('restore presentation subject');
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(model.subject).toBe('');

    model.subject = undefined;
    expect(model.subject).toBeUndefined();
    updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).not.toContain('<d:subject');
    expect(updated).toContain('<d:title>Quarterly</d:title>');
    expect(updated).toContain('<d:creator>Alice</d:creator>');
    expect(updated).toContain('<cp:lastModifiedBy>Editor</cp:lastModifiedBy>');
    expect(updated).toContain('<cp:revision>7</cp:revision>');
    expect(updated).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    expect(pkg.hasPart('/metadata/properties.xml')).toBe(true);
    expect(pkg.relationships('/').filter(
      ({ type }) => type === CORE_PROPERTIES_RELATIONSHIP,
    )).toHaveLength(1);

    model.subject = 'Reopened subject';
    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    expect(reopened.subject).toBe('Reopened subject');
    expect(reopened.slides.map(({ partUri }) => partUri)).toEqual(
      model.slides.map(({ partUri }) => partUri),
    );
    const reopenedXml = new TextDecoder().decode(
      reopened.opcPackage.requirePart('/metadata/properties.xml').bytes,
    );
    expect(reopenedXml).toContain('<d:title>Quarterly</d:title>');
    expect(reopenedXml).toContain('<d:creator>Alice</d:creator>');

    const missingPkg = await OpcPackage.open(await modelFixture());
    const missingModel = new PresentationModel(missingPkg);
    const missingSlide = missingModel.slides[0];
    const beforeAbsentClear = missingPkg.parts.map(({ uri, bytes }) => [uri, bytes] as const);
    const absentJournal = [...missingPkg.mutations];
    expect(missingModel.subject).toBeUndefined();
    missingModel.subject = undefined;
    expect(missingPkg.mutations).toEqual(absentJournal);
    for (const [uri, bytes] of beforeAbsentClear) {
      expect(missingPkg.requirePart(uri).bytes).toEqual(bytes);
    }

    missingModel.subject = 'Created metadata';
    expect(missingModel.subject).toBe('Created metadata');
    expect(missingModel.slides[0]).toBe(missingSlide);
    const createdRelationship = missingPkg.relationships('/').find(
      ({ type }) => type === CORE_PROPERTIES_RELATIONSHIP,
    );
    expect(createdRelationship?.resolvedTarget).toBe('/docProps/core.xml');
    expect(missingPkg.requirePart('/docProps/core.xml').contentType)
      .toBe(CORE_PROPERTIES_CONTENT_TYPE);
    const createdXml = new TextDecoder().decode(
      missingPkg.requirePart('/docProps/core.xml').bytes,
    );
    expect(createdXml).toContain('<dc:subject>Created metadata</dc:subject>');
    expect(createdXml).not.toContain('creator');
    const reopenedCreated = new PresentationModel(await OpcPackage.open(await missingPkg.write()));
    expect(reopenedCreated.subject).toBe('Created metadata');
  });

  it('reads, edits, clears, repairs, creates, and reopens presentation revision metadata', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const coreXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:coreProperties xmlns:c="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
      'xmlns:d="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
      'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      '<d:title>Quarterly</d:title><d:subject>Forecast</d:subject>' +
      '<d:creator>Alice</d:creator><c:lastModifiedBy>Editor</c:lastModifiedBy>' +
      '<c:revision>007</c:revision>' +
      '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>' +
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-30T01:00:00Z</dcterms:modified>' +
      '<x:opaque xmlns:x="urn:test">KEEP</x:opaque></c:coreProperties>';
    pkg.transaction(() => {
      pkg.setPart('/metadata/properties.xml', coreXml, CORE_PROPERTIES_CONTENT_TYPE);
      pkg.addRelationship('/', {
        type: CORE_PROPERTIES_RELATIONSHIP,
        target: 'metadata/properties.xml',
      });
    });
    const model = new PresentationModel(pkg);
    const slide = model.slides[0];
    const readJournal = [...pkg.mutations];
    expect(model.revision).toBe('007');
    expect(pkg.mutations).toEqual(readJournal);

    const beforeSame = pkg.requirePart('/metadata/properties.xml').bytes;
    model.revision = '007';
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeSame);
    expect(pkg.mutations).toEqual(readJournal);
    expect(model.slides[0]).toBe(slide);

    const beforeInvalid = pkg.requirePart('/metadata/properties.xml').bytes;
    const invalidJournal = [...pkg.mutations];
    for (const value of [
      '', ' ', '+1', '-1', '1.0', '1e3', '１２', null, false, 0, 1n, {}, [],
      Symbol('revision'),
    ]) {
      expect(() => {
        model.revision = value as never;
      }).toThrow(TypeError);
    }
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);
    expect(model.slides[0]).toBe(slide);

    const otherParts = new Map(
      pkg.parts
        .filter(({ uri }) => uri !== '/metadata/properties.xml')
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    model.revision = '42';
    expect(model.revision).toBe('42');
    let updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).toContain('<c:revision>42</c:revision>');
    expect(updated).toContain('<d:title>Quarterly</d:title>');
    expect(updated).toContain('<d:subject>Forecast</d:subject>');
    expect(updated).toContain('<d:creator>Alice</d:creator>');
    expect(updated).toContain('<c:lastModifiedBy>Editor</c:lastModifiedBy>');
    expect(updated).toContain('2026-07-30T00:00:00Z');
    expect(updated).toContain('2026-07-30T01:00:00Z');
    expect(updated).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    for (const [uri, bytes] of otherParts) {
      expect(pkg.requirePart(uri).bytes).toEqual(bytes);
    }
    expect(model.slides[0]).toBe(slide);

    model.revision = '0009';
    expect(model.revision).toBe('0009');
    updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).toContain('<c:revision>0009</c:revision>');

    const beforeRollback = pkg.requirePart('/metadata/properties.xml').bytes;
    const rollbackJournal = [...pkg.mutations];
    expect(() => pkg.transaction(() => {
      model.revision = '99';
      expect(model.revision).toBe('99');
      throw new Error('restore presentation revision');
    })).toThrow('restore presentation revision');
    expect(pkg.requirePart('/metadata/properties.xml').bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(model.revision).toBe('0009');

    model.revision = undefined;
    expect(model.revision).toBeUndefined();
    updated = new TextDecoder().decode(pkg.requirePart('/metadata/properties.xml').bytes);
    expect(updated).not.toContain('<c:revision');
    expect(updated).toContain('<d:title>Quarterly</d:title>');
    expect(updated).toContain('<d:subject>Forecast</d:subject>');
    expect(updated).toContain('<d:creator>Alice</d:creator>');
    expect(updated).toContain('<c:lastModifiedBy>Editor</c:lastModifiedBy>');
    expect(updated).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    expect(pkg.hasPart('/metadata/properties.xml')).toBe(true);

    model.revision = '12345678901234567890';
    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    expect(reopened.revision).toBe('12345678901234567890');
    expect(reopened.slides.map(({ partUri }) => partUri)).toEqual(
      model.slides.map(({ partUri }) => partUri),
    );

    const invalidPkg = await OpcPackage.open(await modelFixture());
    invalidPkg.transaction(() => {
      invalidPkg.setPart(
        '/metadata/invalid-core.xml',
        '<?xml version="1.0"?><c:coreProperties ' +
          'xmlns:c="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
          'xmlns:d="http://purl.org/dc/elements/1.1/"><c:revision>abc</c:revision>' +
          '<d:title>KEEP</d:title></c:coreProperties>',
        CORE_PROPERTIES_CONTENT_TYPE,
      );
      invalidPkg.addRelationship('/', {
        type: CORE_PROPERTIES_RELATIONSHIP,
        target: 'metadata/invalid-core.xml',
      });
    });
    const invalidModel = new PresentationModel(invalidPkg);
    expect(invalidModel.revision).toBeUndefined();
    invalidModel.revision = '8';
    expect(invalidModel.revision).toBe('8');
    expect(new TextDecoder().decode(
      invalidPkg.requirePart('/metadata/invalid-core.xml').bytes,
    )).toContain('<c:revision>8</c:revision><d:title>KEEP</d:title>');
    invalidModel.revision = undefined;
    expect(invalidModel.revision).toBeUndefined();

    const missingPkg = await OpcPackage.open(await modelFixture());
    const missingModel = new PresentationModel(missingPkg);
    const missingSlide = missingModel.slides[0];
    const beforeAbsentClear = missingPkg.parts.map(({ uri, bytes }) => [uri, bytes] as const);
    const absentJournal = [...missingPkg.mutations];
    expect(missingModel.revision).toBeUndefined();
    missingModel.revision = undefined;
    expect(missingPkg.mutations).toEqual(absentJournal);
    for (const [uri, bytes] of beforeAbsentClear) {
      expect(missingPkg.requirePart(uri).bytes).toEqual(bytes);
    }

    missingModel.revision = '7';
    expect(missingModel.revision).toBe('7');
    expect(missingModel.slides[0]).toBe(missingSlide);
    const createdRelationship = missingPkg.relationships('/').find(
      ({ type }) => type === CORE_PROPERTIES_RELATIONSHIP,
    );
    expect(createdRelationship?.resolvedTarget).toBe('/docProps/core.xml');
    expect(missingPkg.requirePart('/docProps/core.xml').contentType)
      .toBe(CORE_PROPERTIES_CONTENT_TYPE);
    const createdXml = new TextDecoder().decode(
      missingPkg.requirePart('/docProps/core.xml').bytes,
    );
    expect(createdXml).toContain('<cp:revision>7</cp:revision>');
    expect(createdXml.match(/xmlns:cp=/g)).toHaveLength(1);
    expect(createdXml).not.toContain('creator');
    const reopenedCreated = new PresentationModel(await OpcPackage.open(await missingPkg.write()));
    expect(reopenedCreated.revision).toBe('7');
  });

  it('reads, edits, clears, creates, and reopens presentation company metadata', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const appXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<ep:Properties xmlns:ep="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
      'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
      '<ep:Application>Microsoft Office PowerPoint</ep:Application>' +
      '<ep:PresentationFormat>Custom</ep:PresentationFormat><ep:Slides>2</ep:Slides>' +
      '<ep:HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant>' +
      '<vt:lpstr>Theme</vt:lpstr></vt:variant></vt:vector></ep:HeadingPairs>' +
      '<ep:Company custom="KEEP">Original &amp; company</ep:Company>' +
      '<x:opaque xmlns:x="urn:test">KEEP</x:opaque>' +
      '<ep:AppVersion>16.0000</ep:AppVersion></ep:Properties>';
    pkg.transaction(() => {
      pkg.setPart(
        '/metadata/application.xml',
        appXml,
        EXTENDED_PROPERTIES_CONTENT_TYPE,
      );
      pkg.addRelationship('/', {
        type: EXTENDED_PROPERTIES_RELATIONSHIP,
        target: 'metadata/application.xml',
      });
    });
    const model = new PresentationModel(pkg);
    const slide = model.slides[0];
    const readJournal = [...pkg.mutations];
    expect(model.company).toBe('Original & company');
    expect(pkg.mutations).toEqual(readJournal);

    const beforeSame = pkg.requirePart('/metadata/application.xml').bytes;
    model.company = 'Original & company';
    expect(pkg.requirePart('/metadata/application.xml').bytes).toEqual(beforeSame);
    expect(pkg.mutations).toEqual(readJournal);
    expect(model.slides[0]).toBe(slide);

    const beforeInvalid = pkg.requirePart('/metadata/application.xml').bytes;
    const invalidJournal = [...pkg.mutations];
    for (const value of [
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
      expect(() => {
        model.company = value as never;
      }).toThrow(TypeError);
    }
    expect(pkg.requirePart('/metadata/application.xml').bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);
    expect(model.slides[0]).toBe(slide);

    const otherParts = new Map(
      pkg.parts
        .filter(({ uri }) => uri !== '/metadata/application.xml')
        .map(({ uri, bytes }) => [uri, bytes]),
    );
    model.company = 'Edited & <safe>';
    expect(model.company).toBe('Edited & <safe>');
    let updated = new TextDecoder().decode(
      pkg.requirePart('/metadata/application.xml').bytes,
    );
    expect(updated).toContain(
      '<ep:Company custom="KEEP">Edited &amp; &lt;safe&gt;</ep:Company>',
    );
    expect(updated).toContain(
      '<ep:Application>Microsoft Office PowerPoint</ep:Application>',
    );
    expect(updated).toContain('<ep:PresentationFormat>Custom</ep:PresentationFormat>');
    expect(updated).toContain('<ep:Slides>2</ep:Slides>');
    expect(updated).toContain('<vt:lpstr>Theme</vt:lpstr>');
    expect(updated).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    expect(updated).toContain('<ep:AppVersion>16.0000</ep:AppVersion>');
    for (const [uri, bytes] of otherParts) {
      expect(pkg.requirePart(uri).bytes).toEqual(bytes);
    }
    expect(model.slides[0]).toBe(slide);

    model.company = '';
    expect(model.company).toBe('');
    updated = new TextDecoder().decode(pkg.requirePart('/metadata/application.xml').bytes);
    expect(updated).toContain('<ep:Company custom="KEEP"></ep:Company>');

    const beforeRollback = pkg.requirePart('/metadata/application.xml').bytes;
    const rollbackJournal = [...pkg.mutations];
    expect(() => pkg.transaction(() => {
      model.company = 'Temporary';
      expect(model.company).toBe('Temporary');
      throw new Error('restore presentation company');
    })).toThrow('restore presentation company');
    expect(pkg.requirePart('/metadata/application.xml').bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(model.company).toBe('');

    model.company = undefined;
    expect(model.company).toBeUndefined();
    updated = new TextDecoder().decode(pkg.requirePart('/metadata/application.xml').bytes);
    expect(updated).not.toContain('<ep:Company');
    expect(updated).toContain(
      '<ep:Application>Microsoft Office PowerPoint</ep:Application>',
    );
    expect(updated).toContain('<ep:AppVersion>16.0000</ep:AppVersion>');
    expect(pkg.hasPart('/metadata/application.xml')).toBe(true);
    expect(pkg.relationships('/').filter(
      ({ type }) => type === EXTENDED_PROPERTIES_RELATIONSHIP,
    )).toHaveLength(1);

    model.company = 'Reopened company';
    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    expect(reopened.company).toBe('Reopened company');
    expect(reopened.slides.map(({ partUri }) => partUri)).toEqual(
      model.slides.map(({ partUri }) => partUri),
    );
    expect(new TextDecoder().decode(
      reopened.opcPackage.requirePart('/metadata/application.xml').bytes,
    )).toContain('<ep:AppVersion>16.0000</ep:AppVersion>');

    const missingPkg = await OpcPackage.open(await modelFixture());
    const missingModel = new PresentationModel(missingPkg);
    const missingSlide = missingModel.slides[0];
    const beforeAbsentClear = missingPkg.parts.map(({ uri, bytes }) => [uri, bytes] as const);
    const absentJournal = [...missingPkg.mutations];
    expect(missingModel.company).toBeUndefined();
    missingModel.company = undefined;
    expect(missingPkg.mutations).toEqual(absentJournal);
    for (const [uri, bytes] of beforeAbsentClear) {
      expect(missingPkg.requirePart(uri).bytes).toEqual(bytes);
    }

    missingModel.company = 'Created metadata';
    expect(missingModel.company).toBe('Created metadata');
    expect(missingModel.slides[0]).toBe(missingSlide);
    const createdRelationship = missingPkg.relationships('/').find(
      ({ type }) => type === EXTENDED_PROPERTIES_RELATIONSHIP,
    );
    expect(createdRelationship?.resolvedTarget).toBe('/docProps/app.xml');
    expect(missingPkg.requirePart('/docProps/app.xml').contentType)
      .toBe(EXTENDED_PROPERTIES_CONTENT_TYPE);
    const createdXml = new TextDecoder().decode(
      missingPkg.requirePart('/docProps/app.xml').bytes,
    );
    expect(createdXml).toContain('<Company>Created metadata</Company>');
    expect(createdXml).not.toContain('Application');
    const reopenedCreated = new PresentationModel(await OpcPackage.open(await missingPkg.write()));
    expect(reopenedCreated.company).toBe('Created metadata');
  });

  it('uses r:id order and exposes common semantic objects', async () => {
    const model = new PresentationModel(await OpcPackage.open(await modelFixture()));
    const initialSlides = model.slides;
    expect(model.slides[0]).toBe(initialSlides[0]);
    expect(model.slides.map(({ title }) => title.text)).toEqual(['Second title', 'First title']);
    const shapes = model.slides[1]!.shapes;
    expect(model.slides[1]!.shapes[0]).toBe(shapes[0]);
    expect(model.slides[1]!.shapes[1]).toBe(shapes[1]);
    expect(shapes[0]).toBeInstanceOf(ShapeModel);
    expect(shapes[1]).toBeInstanceOf(ImageModel);
    expect(shapes[2]).toBeInstanceOf(TableModel);
    expect(shapes[3]).toBeInstanceOf(ChartModel);
    expect(emuToInches(shapes[0]!.transform.x)).toBe(1);
    expect((shapes[1] as ImageModel).sourcePartUri).toBe('/ppt/media/image1.png');
    expect((shapes[2] as TableModel).rows[0]?.cells.map(({ text }) => text)).toEqual(['A1', 'B1']);
    expect((shapes[3] as ChartModel).chartPartUri).toBe('/ppt/charts/chart1.xml');
    expect((shapes[3] as ChartModel).workbookPartUri).toBe('/ppt/embeddings/workbook1.xlsx');
    expect((shapes[3] as ChartModel).definition).toEqual({
      groups: [{
        type: 'bar',
        axis: 'primary',
        series: [{ name: 'Sales', categories: ['Q1', 'Q2'], values: [10, 20] }],
      }],
      options: {},
    });
    expect((shapes[3] as ChartModel).series).toEqual([
      { name: 'Sales', categories: ['Q1', 'Q2'], values: [10, 20] },
    ]);
    (shapes[0] as ShapeModel).setTransform({ x: inches(2) });
    expect(model.slides[1]!.shapes[0]).toBe(shapes[0]);
    expect(emuToInches(model.slides[1]!.shapes[0]!.transform.x)).toBe(2);
    (shapes[2] as TableModel).setCellText(0, 1, 'Edited B1');
    expect(model.slides[1]!.shapes[2]).toBe(shapes[2]);
    expect((model.slides[1]!.shapes[2] as TableModel).rows[0]?.cells[1]?.text).toBe('Edited B1');
    const imagePartUri = (shapes[1] as ImageModel).sourcePartUri;
    (shapes[1] as ImageModel).replaceData(new Uint8Array([1, 2, 3]), 'image/png');
    expect((shapes[1] as ImageModel).sourcePartUri).toBe(imagePartUri);
    expect(model.opcPackage.requirePart('/ppt/media/image1.png').bytes).toEqual(new Uint8Array([1, 2, 3]));
    const chartPartUri = (shapes[3] as ChartModel).chartPartUri;
    (shapes[3] as ChartModel).setXml('<c:chartSpace xmlns:c="c"><c:chart><c:plotArea/></c:chart></c:chartSpace>');
    expect((shapes[3] as ChartModel).chartPartUri).toBe(chartPartUri);
    expect((model.slides[1]!.shapes[3] as ChartModel).xml).toContain('<c:plotArea/>');
  });

  it('creates and reopens all categorical and axis-free chart types atomically', async () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const slidePart = pkg.requirePart(slide.partUri);
    pkg.setPart(
      slide.partUri,
      new TextDecoder().decode(slidePart.bytes).replace(
        '</p:spTree>',
        '<p:extLst><p:ext uri="urn:test"><x:opaque xmlns:x="urn:test">KEEP</x:opaque>'
          + '</p:ext></p:extLst></p:spTree>',
      ),
      slidePart.contentType,
    );
    const types = ['area', 'bar', 'bar3D', 'doughnut', 'line', 'pie', 'radar'] as const;
    const created: ChartModel[] = [];
    for (const [index, type] of types.entries()) {
      const chart = await slide.addChart(type, [{
        name: `${type} series`,
        categories: ['Q1', 'Q2'],
        values: [10, 20],
      }], index === 0 ? {
        name: 'Revenue & Cost',
        altText: 'Quarterly chart',
        x: inches(2),
        y: inches(1.5),
        width: inches(7),
        height: inches(3),
        rotation: degrees(15),
        flipHorizontal: true,
      } : {});
      created.push(chart);
      expect(slide.shapes.find(({ id }) => id === chart.id)).toBe(chart);
      expect(chart.definition?.groups[0]?.type).toBe(type);
      expect(chart.series).toEqual([{
        name: `${type} series`,
        categories: ['Q1', 'Q2'],
        values: [10, 20],
      }]);
      expect(chart.workbookPartUri).toMatch(/^\/ppt\/embeddings\/Microsoft_Excel_Worksheet\d+\.xlsx$/);
      expect(pkg.relationships(chart.chartPartUri!)).toMatchObject([{
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package',
        targetMode: 'Internal',
        resolvedTarget: chart.workbookPartUri,
      }]);
    }
    expect(created[0]?.name).toBe('Revenue & Cost');
    expect(created[0]?.altText).toBe('Quarterly chart');
    expect(created[0]?.transform).toMatchObject({
      x: inches(2),
      y: inches(1.5),
      width: inches(7),
      height: inches(3),
      rotation: degrees(15),
      flipHorizontal: true,
      flipVertical: false,
    });
    const slideXml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(slideXml.indexOf('name="Chart"')).toBeLessThan(slideXml.indexOf('<p:extLst>'));
    expect(slideXml).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    expect(pkg.parts.filter(({ contentType }) =>
      contentType === 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'))
      .toHaveLength(7);
    expect(pkg.parts.filter(({ contentType }) =>
      contentType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'))
      .toHaveLength(7);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedCharts = reopened.slides[0]!.shapes.filter(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    );
    expect(reopenedCharts.map((chart) => chart.definition?.groups[0]?.type)).toEqual(types);
    expect(reopenedCharts.map((chart) => chart.series[0]?.values)).toEqual(
      types.map(() => [10, 20]),
    );
  });

  it('rolls back every native chart creation failure boundary', async () => {
    const stages = [
      'workbook part',
      'chart part',
      'workbook relationship',
      'slide relationship',
      'slide XML',
      'outer transaction',
    ] as const;
    for (const stage of stages) {
      const { pkg, model } = emptyPresentationModel();
      const slide = model.addSlide();
      const before = packageSnapshot(pkg);
      const restores: (() => void)[] = [];
      if (stage === 'workbook part' || stage === 'chart part' || stage === 'slide XML') {
        const failAt = stage === 'workbook part' ? 1 : stage === 'chart part' ? 2 : 3;
        const original = pkg.setPart.bind(pkg);
        let calls = 0;
        const spy = vi.spyOn(pkg, 'setPart').mockImplementation((uri, bytes, contentType) => {
          calls += 1;
          if (calls === failAt) throw new Error(`injected ${stage}`);
          return original(uri, bytes, contentType);
        });
        restores.push(() => spy.mockRestore());
      } else if (stage === 'workbook relationship' || stage === 'slide relationship') {
        const failAt = stage === 'workbook relationship' ? 1 : 2;
        const original = pkg.addRelationship.bind(pkg);
        let calls = 0;
        const spy = vi.spyOn(pkg, 'addRelationship').mockImplementation((sourceUri, input) => {
          calls += 1;
          if (calls === failAt) throw new Error(`injected ${stage}`);
          return original(sourceUri, input);
        });
        restores.push(() => spy.mockRestore());
      } else {
        const original = pkg.transaction.bind(pkg);
        const spy = vi.spyOn(pkg, 'transaction').mockImplementation(((operation: () => unknown) =>
          original(() => {
            operation();
            throw new Error('injected outer transaction');
          })) as typeof pkg.transaction);
        restores.push(() => spy.mockRestore());
      }
      try {
        await expect(slide.addChart('bar', [{
          name: 'Revenue',
          categories: ['Q1', 'Q2'],
          values: [10, 20],
        }])).rejects.toThrow(`injected ${stage}`);
        expect(packageSnapshot(pkg)).toEqual(before);
      } finally {
        restores.forEach((restore) => restore());
      }
    }
  });

  it('creates, duplicates, and reopens independent scatter and bubble vectors', async () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const scatter = await slide.addChart('scatter', [
      { name: 'First', xValues: [0, 2, 4], values: [1, 3, 5] },
      { name: 'Second', xValues: [10, 20, 30], values: [11, 21, 31] },
    ]);
    const bubble = await slide.addChart('bubble', [{
      name: 'Bubbles',
      xValues: [1, 2, 3],
      values: [4, 5, 6],
      sizes: [7, 8, 9],
    }]);
    expect(scatter.series).toEqual([
      { name: 'First', values: [1, 3, 5], xValues: [0, 2, 4] },
      { name: 'Second', values: [11, 21, 31], xValues: [10, 20, 30] },
    ]);
    expect(bubble.series).toEqual([{
      name: 'Bubbles', values: [4, 5, 6], xValues: [1, 2, 3], sizes: [7, 8, 9],
    }]);
    expect(await chartWorkbookMatches(
      pkg.requirePart(scatter.workbookPartUri!).bytes,
      scatter.definition!,
    )).toBe(true);
    expect(await chartWorkbookMatches(
      pkg.requirePart(bubble.workbookPartUri!).bytes,
      bubble.definition!,
    )).toBe(true);

    const duplicate = model.duplicateSlide(0);
    const duplicateCharts = duplicate.shapes.filter(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    );
    expect(duplicateCharts.map(({ series }) => series)).toEqual([scatter.series, bubble.series]);
    expect(duplicateCharts[0]?.chartPartUri).not.toBe(scatter.chartPartUri);
    expect(duplicateCharts[0]?.workbookPartUri).not.toBe(scatter.workbookPartUri);
    expect(duplicateCharts[1]?.chartPartUri).not.toBe(bubble.chartPartUri);
    expect(duplicateCharts[1]?.workbookPartUri).not.toBe(bubble.workbookPartUri);

    const invalid = [
      ['scatter', [{ name: 'Bad', xValues: [1], values: [Number.NaN] }]],
      ['scatter', [{ name: 'Bad', xValues: [1], values: [Number.POSITIVE_INFINITY] }]],
      ['scatter', [{ name: 'Bad', xValues: [1, 2], values: [3] }]],
      ['scatter', [{ name: 'Bad', xValues: [], values: [] }]],
      ['scatter', [{ name: 'Bad', categories: ['A'], xValues: [1], values: [2] }]],
      ['bubble', [{ name: 'Bad', xValues: [1], values: [2], sizes: [0] }]],
      ['bubble', [{ name: 'Bad', xValues: [1], values: [2], sizes: [-1] }]],
    ] as const;
    for (const [type, series] of invalid) {
      const before = packageSnapshot(pkg);
      await expect(slide.addChart(type, series as never)).rejects.toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
    }
    const valid = await slide.addChart('scatter', [{
      name: 'Still next', xValues: [1], values: [2],
    }]);
    expect(valid.id).toBe(bubble.id + 1);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedCharts = reopened.slides.flatMap(({ shapes }) => shapes).filter(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    );
    expect(reopenedCharts.filter((chart) => chart.definition?.groups[0]?.type === 'scatter'))
      .toHaveLength(3);
    expect(reopenedCharts.filter((chart) => chart.definition?.groups[0]?.type === 'bubble'))
      .toHaveLength(2);
  });

  it('creates, duplicates, and reopens ordered combination groups and secondary axes', async () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const combo = await slide.addChart([
      {
        type: 'bar',
        series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
      },
      {
        type: 'line',
        axis: 'secondary',
        series: [{ name: 'Trend', categories: ['Q1', 'Q2'], values: [11, 21] }],
      },
    ], { name: 'Revenue combo' });
    expect(combo.name).toBe('Revenue combo');
    expect(combo.definition?.groups).toEqual([
      {
        type: 'bar',
        axis: 'primary',
        series: [{ name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20] }],
      },
      {
        type: 'line',
        axis: 'secondary',
        series: [{ name: 'Trend', categories: ['Q1', 'Q2'], values: [11, 21] }],
      },
    ]);
    expect(combo.series.map(({ name }) => name)).toEqual(['Revenue', 'Trend']);
    expect(combo.xml.indexOf('<c:barChart>')).toBeLessThan(combo.xml.indexOf('<c:lineChart>'));
    expect(combo.xml).toContain('<c:axId val="10000003"/><c:axId val="10000004"/>');
    expect(await chartWorkbookMatches(
      pkg.requirePart(combo.workbookPartUri!).bytes,
      combo.definition!,
    )).toBe(true);

    const primaryOnly = await slide.addChart([
      {
        type: 'area',
        series: [{ name: 'Actual', categories: ['Q1'], values: [10] }],
      },
      {
        type: 'line',
        series: [{ name: 'Plan', categories: ['Q1'], values: [11] }],
      },
      {
        type: 'line',
        series: [{ name: 'Forecast', categories: ['Q1'], values: [12] }],
      },
    ]);
    expect(primaryOnly.definition?.groups.map(({ type, axis }) => [type, axis])).toEqual([
      ['area', 'primary'],
      ['line', 'primary'],
      ['line', 'primary'],
    ]);
    expect(primaryOnly.xml).not.toContain('10000003');

    const invalidGroups = [
      [],
      [
        { type: 'bar', series: [{ name: 'A', categories: ['Q1'], values: [1] }] },
        { type: 'pie', series: [{ name: 'B', categories: ['Q1'], values: [2] }] },
      ],
      [
        { type: 'bar3D', series: [{ name: 'A', categories: ['Q1'], values: [1] }] },
        { type: 'line', series: [{ name: 'B', categories: ['Q1'], values: [2] }] },
      ],
      [
        { type: 'bubble', series: [{ name: 'A', xValues: [1], values: [2], sizes: [3] }] },
        { type: 'scatter', series: [{ name: 'B', xValues: [1], values: [2] }] },
      ],
      [
        { type: 'scatter', series: [{ name: 'A', xValues: [1], values: [2] }] },
        { type: 'line', series: [{ name: 'B', categories: ['Q1'], values: [2] }] },
      ],
      [
        {
          type: 'line',
          axis: 'secondary',
          series: [{ name: 'A', categories: ['Q1'], values: [1] }],
        },
        { type: 'bar', series: [{ name: 'B', categories: ['Q1'], values: [2] }] },
      ],
    ];
    for (const groups of invalidGroups) {
      const before = packageSnapshot(pkg);
      await expect(slide.addChart(groups as never)).rejects.toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    const duplicate = model.duplicateSlide(0);
    const duplicateCharts = duplicate.shapes.filter(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    );
    expect(duplicateCharts.map((chart) => chart.definition?.groups.map(({ type }) => type)))
      .toEqual([['bar', 'line'], ['area', 'line', 'line']]);
    expect(duplicateCharts[0]?.chartPartUri).not.toBe(combo.chartPartUri);
    expect(duplicateCharts[0]?.workbookPartUri).not.toBe(combo.workbookPartUri);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedCombo = reopened.slides[0]!.shapes.find(
      (shape): shape is ChartModel => shape instanceof ChartModel && shape.name === 'Revenue combo',
    );
    expect(reopenedCombo?.definition?.groups.map(({ type, axis }) => [type, axis])).toEqual([
      ['bar', 'primary'],
      ['line', 'secondary'],
    ]);
  });

  it('edits and reopens synchronized chart semantics through the live model', async () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const chart = await slide.addChart('bar', [{
      name: 'Revenue', categories: ['Q1', 'Q2'], values: [10, 20],
    }]);
    const identity = chart;
    const invalidBefore = packageSnapshot(pkg);
    await expect(chart.replaceDefinition({
      groups: [{
        type: 'bar',
        series: [{
          name: 'Invalid', categories: ['Q1'], values: [1], extra: true,
        }],
      }],
    } as never)).rejects.toThrow(/unsupported property/);
    expect(packageSnapshot(pkg)).toEqual(invalidBefore);

    await chart.replaceSeries([{
      name: 'Revenue edited', categories: ['Q1', 'Q2', 'Q3'], values: [12, 24, 36],
    }]);
    expect(slide.shapes.find(({ id }) => id === chart.id)).toBe(identity);
    expect(chart.series[0]?.values).toEqual([12, 24, 36]);
    expect(await chartWorkbookMatches(
      pkg.requirePart(chart.workbookPartUri!).bytes,
      chart.definition!,
    )).toBe(true);

    await chart.replaceDefinition({ groups: [{
      type: 'scatter',
      series: [{ name: 'Points', xValues: [1, 2], values: [3, 4] }],
    }] });
    expect(chart.definition?.groups[0]?.type).toBe('scatter');
    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedChart = reopened.slides[0]!.shapes.find(
      (shape): shape is ChartModel => shape instanceof ChartModel,
    )!;
    expect(reopenedChart.definition).toEqual(chart.definition);
    expect(await chartWorkbookMatches(
      reopened.opcPackage.requirePart(reopenedChart.workbookPartUri!).bytes,
      reopenedChart.definition!,
    )).toBe(true);
  });

  it('deletes exclusive and shared chart dependency roots atomically without renumbering peers', async () => {
    const { pkg, model } = emptyPresentationModel();
    const slide = model.addSlide();
    const source = await slide.addChart('bar', [{
      name: 'Revenue', categories: ['Q1'], values: [10],
    }]);
    const peer = await slide.addChart('line', [{
      name: 'Trend', categories: ['Q1'], values: [11],
    }]);
    const sourceChartUri = source.chartPartUri!;
    const sourceWorkbookUri = source.workbookPartUri!;
    pkg.setPart('/ppt/charts/style1.xml', '<style/>', 'application/vnd.test.chart-style+xml');
    pkg.setPart('/ppt/charts/colors1.xml', '<colors/>', 'application/vnd.test.chart-colors+xml');
    pkg.setPart('/ppt/media/shared-chart.png', Uint8Array.of(137, 80, 78, 71), 'image/png');
    pkg.addRelationship(sourceChartUri, {
      type: 'http://schemas.microsoft.com/office/2011/relationships/chartStyle',
      target: 'style1.xml',
    });
    pkg.addRelationship('/ppt/charts/style1.xml', {
      type: 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle',
      target: 'colors1.xml',
    });
    pkg.addRelationship(sourceChartUri, {
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      target: '../media/shared-chart.png',
    });
    pkg.addRelationship(sourceChartUri, {
      type: 'https://example.com/relationships/externalData',
      target: 'https://example.com/chart-data',
      targetMode: 'External',
    });

    const duplicateSlide = model.duplicateSlide(0);
    const duplicate = duplicateSlide.shapes.find(
      (shape): shape is ChartModel => shape instanceof ChartModel && shape.name === source.name,
    )!;
    const duplicateChartUri = duplicate.chartPartUri!;
    const duplicateWorkbookUri = duplicate.workbookPartUri!;
    const duplicateRelationship = duplicateSlide.relationships.find(
      ({ resolvedTarget }) => resolvedTarget === duplicateChartUri,
    )!;
    pkg.updateRelationship(duplicateSlide.partUri, duplicateRelationship.id, {
      target: relativeRelationshipTarget(duplicateSlide.partUri, sourceChartUri),
    });
    pkg.deletePart(duplicateChartUri);
    pkg.deletePart(duplicateWorkbookUri);

    const peerIdentity = peer;
    const nextExpectedId = Math.max(...slide.shapes.map(({ id }) => id)) + 1;
    source.remove();
    expect(slide.shapes.find(({ id }) => id === peer.id)).toBe(peerIdentity);
    expect(pkg.hasPart(sourceChartUri)).toBe(true);
    expect(pkg.hasPart(sourceWorkbookUri)).toBe(true);
    expect(pkg.hasPart('/ppt/charts/style1.xml')).toBe(true);
    expect(pkg.hasPart('/ppt/charts/colors1.xml')).toBe(true);
    expect(pkg.hasPart('/ppt/media/shared-chart.png')).toBe(true);

    const next = await slide.addChart('pie', [{
      name: 'Share', categories: ['A'], values: [1],
    }]);
    expect(next.id).toBe(nextExpectedId);

    duplicate.remove();
    expect(pkg.hasPart(sourceChartUri)).toBe(false);
    expect(pkg.hasPart(sourceWorkbookUri)).toBe(false);
    expect(pkg.hasPart('/ppt/charts/style1.xml')).toBe(false);
    expect(pkg.hasPart('/ppt/charts/colors1.xml')).toBe(false);
    expect(pkg.hasPart('/ppt/media/shared-chart.png')).toBe(true);

    const rollback = await slide.addChart('bar', [{
      name: 'Rollback', categories: ['Q1'], values: [1],
    }]);
    const before = packageSnapshot(pkg);
    const rollbackUri = rollback.chartPartUri!;
    const originalDelete = pkg.deletePart.bind(pkg);
    const spy = vi.spyOn(pkg, 'deletePart').mockImplementation((uri) => {
      if (uri === rollbackUri) throw new Error('injected chart delete');
      originalDelete(uri);
    });
    try {
      expect(() => rollback.remove()).toThrow('injected chart delete');
      expect(packageSnapshot(pkg)).toEqual(before);
      expect(slide.shapes.find(({ id }) => id === rollback.id)).toBe(rollback);
    } finally {
      spy.mockRestore();
    }
  });

  it('deletes imported chart roots and reopens every presentation format', async () => {
    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const model = new PresentationModel(await OpcPackage.open(
        await modelFixture(profile.presentationContentType),
      ));
      const slide = model.slides[1]!;
      const chart = slide.shapes.find(
        (shape): shape is ChartModel => shape instanceof ChartModel,
      )!;
      chart.remove();
      expect(slide.shapes.some((shape) => shape instanceof ChartModel)).toBe(false);
      expect(model.opcPackage.hasPart('/ppt/charts/chart1.xml')).toBe(false);
      expect(model.opcPackage.hasPart('/ppt/embeddings/workbook1.xlsx')).toBe(false);
      const reopened = new PresentationModel(await OpcPackage.open(await model.opcPackage.write()));
      expect(reopened.format).toBe(profile.format);
      expect(reopened.slides.flatMap(({ shapes }) => shapes)
        .some((shape) => shape instanceof ChartModel)).toBe(false);
    }
  });

  it('creates table-cell hyperlinks with independent atomic relationship ownership', async () => {
    const { pkg, model } = emptyPresentationModel();
    const source = model.addSlide();
    const target = model.addSlide();
    const firstUrl: { url: string; tooltip?: string } = {
      url: 'https://example.com?a=1&b=2',
      tooltip: 'Visit & learn',
    };
    const secondUrl: { url: string; tooltip?: string } = {
      url: 'https://example.com?a=1&b=2',
      tooltip: '',
    };
    const internal: { slide: number; tooltip?: string } = { slide: 2 };
    const self: Hyperlink = { slide: 1, tooltip: '' };
    const table = source.addTable([[
      { text: 'URL one', options: { hyperlink: firstUrl } },
      { text: 'URL two', options: { hyperlink: secondUrl } },
      { text: 'Target', options: { hyperlink: internal } },
      { text: 'Self', options: { hyperlink: self } },
      'Plain',
    ]], { name: 'Linked table' });

    const expected = [
      { url: 'https://example.com?a=1&b=2', tooltip: 'Visit & learn' },
      { url: 'https://example.com?a=1&b=2', tooltip: '' },
      { slide: 2 },
      { slide: 1, tooltip: '' },
      undefined,
    ];
    expect(table.rows[0]!.cells.map(({ hyperlink }) => hyperlink)).toEqual(expected);
    expect(Object.isFrozen(table.rows[0]!.cells[0]!.hyperlink)).toBe(true);
    const hyperlinkRelationships = source.relationships.filter(
      ({ type }) => type === HYPERLINK_RELATIONSHIP,
    );
    expect(hyperlinkRelationships).toHaveLength(2);
    expect(new Set(hyperlinkRelationships.map(({ id }) => id)).size).toBe(2);
    expect(hyperlinkRelationships.map(({ target }) => target)).toEqual([
      'https://example.com?a=1&b=2',
      'https://example.com?a=1&b=2',
    ]);
    expect(source.relationships.filter(({ type }) => type === SLIDE_RELATIONSHIP))
      .toEqual([
        expect.objectContaining({ resolvedTarget: target.partUri }),
        expect.objectContaining({ resolvedTarget: source.partUri }),
      ]);
    const createdXml = new TextDecoder().decode(pkg.requirePart(source.partUri).bytes);
    const clickIds = [...createdXml.matchAll(/<a:hlinkClick\b[^>]*\br:id="([^"]+)"/g)]
      .map((match) => match[1]);
    expect(clickIds).toHaveLength(4);
    expect(new Set(clickIds).size).toBe(4);
    expect(createdXml.match(/\bu="sng"/g)).toHaveLength(4);
    expect(createdXml).toContain('tooltip="Visit &amp; learn"');
    expect(createdXml).toContain('tooltip="" action="ppaction://hlinksldjump"');

    firstUrl.url = 'https://changed.example';
    firstUrl.tooltip = 'Changed';
    secondUrl.url = 'https://changed.example';
    internal.slide = 1;
    expect(table.rows[0]!.cells.map(({ hyperlink }) => hyperlink)).toEqual(expected);
    expect(Object.hasOwn(firstUrl, '_rId')).toBe(false);

    table.setCellText(0, 0, 'URL edited');
    table.setCellFill(0, 0, { kind: 'none' });
    table.setCellMargins(0, 1, { top: 3 });
    table.setCellHorizontalAlignment(0, 2, 'center');
    table.setCellVerticalAlignment(0, 3, 'bottom');
    expect(table.rows[0]!.cells[0]!.hyperlink).toEqual(expected[0]);
    expect(table.rows[0]!.cells.map(({ hyperlink }) => hyperlink)).toEqual(expected);

    model.moveSlide(model.slides.indexOf(target), 0);
    expect(table.rows[0]!.cells[2]!.hyperlink).toEqual({ slide: 1 });
    expect(table.rows[0]!.cells[3]!.hyperlink).toEqual({ slide: 2, tooltip: '' });
    model.moveSlide(0, model.slides.length - 1);
    expect(table.rows[0]!.cells[2]!.hyperlink).toEqual({ slide: 2 });
    expect(table.rows[0]!.cells[3]!.hyperlink).toEqual({ slide: 1, tooltip: '' });

    const duplicate = model.duplicateSlide(model.slides.indexOf(source));
    const duplicateTable = duplicate.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    )!;
    expect(duplicateTable.rows[0]!.cells[0]!.hyperlink).toEqual(expected[0]);
    expect(duplicateTable.rows[0]!.cells[2]!.hyperlink).toEqual({
      slide: model.slides.indexOf(target) + 1,
    });
    expect(duplicateTable.rows[0]!.cells[3]!.hyperlink).toEqual({
      slide: model.slides.indexOf(duplicate) + 1,
      tooltip: '',
    });
    expect(table.rows[0]!.cells[3]!.hyperlink).toEqual({
      slide: model.slides.indexOf(source) + 1,
      tooltip: '',
    });

    model.deleteSlide(model.slides.indexOf(target));
    expect(table.rows[0]!.cells[2]!.hyperlink).toBeUndefined();
    expect(duplicateTable.rows[0]!.cells[2]!.hyperlink).toBeUndefined();
    expect(table.rows[0]!.cells[0]!.hyperlink).toEqual(expected[0]);
    expect(duplicateTable.rows[0]!.cells[0]!.hyperlink).toEqual(expected[0]);
    expect(table.rows[0]!.cells[3]!.hyperlink).toEqual({
      slide: model.slides.indexOf(source) + 1,
      tooltip: '',
    });
    expect(duplicateTable.rows[0]!.cells[3]!.hyperlink).toEqual({
      slide: model.slides.indexOf(duplicate) + 1,
      tooltip: '',
    });
    expect(new TextDecoder().decode(pkg.requirePart(source.partUri).bytes))
      .toContain('u="sng"');

    const beforeInvalid = packageSnapshot(pkg);
    expect(() => source.addTable([[
      { text: 'Prepared first', options: { hyperlink: { url: 'https://first.example' } } },
      { text: 'Invalid later', options: { hyperlink: { slide: 99 } } },
    ]])).toThrow('Table cell 0,1 hyperlink slide 99 is out of range');
    expect(packageSnapshot(pkg)).toEqual(beforeInvalid);

    const originalAddRelationship = pkg.addRelationship.bind(pkg);
    let addedCellLinks = 0;
    const spy = vi.spyOn(pkg, 'addRelationship').mockImplementation((partUri, input) => {
      if (partUri === source.partUri && input.type === HYPERLINK_RELATIONSHIP) {
        addedCellLinks += 1;
        if (addedCellLinks === 2) throw new Error('injected second cell hyperlink');
      }
      return originalAddRelationship(partUri, input);
    });
    try {
      expect(() => source.addTable([[
        { text: 'First', options: { hyperlink: { url: 'https://first.example' } } },
        { text: 'Second', options: { hyperlink: { url: 'https://second.example' } } },
      ]])).toThrow('injected second cell hyperlink');
      expect(packageSnapshot(pkg)).toEqual(beforeInvalid);
    } finally {
      spy.mockRestore();
    }

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSource = reopened.slides.find(({ partUri }) => partUri === source.partUri)!;
    const reopenedTable = reopenedSource.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    )!;
    expect(reopenedTable.rows[0]!.cells.map(({ hyperlink }) => hyperlink)).toEqual([
      expected[0],
      expected[1],
      undefined,
      { slide: reopened.slides.indexOf(reopenedSource) + 1, tooltip: '' },
      undefined,
    ]);
  });

  it('creates rich table-cell paragraphs and positioned hyperlinks atomically', () => {
    const { pkg, model } = emptyPresentationModel();
    const source = model.addSlide();
    const target = model.addSlide();
    const rich = [{
      align: 'right' as const,
      runs: [
        { text: 'Inherited', style: { bold: true } },
        {
          text: ' explicit',
          style: { hyperlink: { url: 'https://run.example', tooltip: 'Run' } },
        },
        {
          text: ' suppressed',
          style: { hyperlink: false as const, italic: true },
        },
      ],
    }, {
      runs: [{
        text: 'Slide',
        style: { hyperlink: { slide: 2, tooltip: '' } },
      }],
    }];
    source.addTable([[
      {
        text: rich,
        options: { hyperlink: { url: 'https://cell.example' } },
      },
      'one\r\ntwo\n',
    ]], { name: 'Rich linked table' });

    const xml = new TextDecoder().decode(pkg.requirePart(source.partUri).bytes);
    expect(xml.match(/<a:p>/g)).toHaveLength(5);
    const clicks = [...xml.matchAll(/<a:hlinkClick\b[^>]*\br:id="([^"]+)"[^>]*>/g)];
    expect(clicks).toHaveLength(3);
    const clickIds = clicks.map((match) => match[1]!);
    expect(new Set(clickIds).size).toBe(3);
    expect(xml).toContain('<a:pPr algn="r"');
    expect(xml).toContain('tooltip="Run"');
    expect(xml).toContain('tooltip="" action="ppaction://hlinksldjump"');
    expect(xml.match(/> suppressed<\/a:t>/g)).toHaveLength(1);

    const byId = new Map(source.relationships.map((relationship) => [
      relationship.id,
      relationship,
    ]));
    expect(byId.get(clickIds[0]!)?.target).toBe('https://cell.example');
    expect(byId.get(clickIds[1]!)?.target).toBe('https://run.example');
    expect(byId.get(clickIds[2]!)?.resolvedTarget).toBe(target.partUri);

    rich[0]!.runs[0]!.text = 'MUTATED';
    expect(new TextDecoder().decode(pkg.requirePart(source.partUri).bytes)).toBe(xml);

    const beforeInvalid = packageSnapshot(pkg);
    expect(() => source.addTable([[
      { text: [{ runs: [{ text: 'Invalid', style: { hyperlink: { slide: 99 } } }] }] },
    ]])).toThrow('Table cell 0,0 hyperlink slide 99 is out of range');
    expect(packageSnapshot(pkg)).toEqual(beforeInvalid);

    const originalAddRelationship = pkg.addRelationship.bind(pkg);
    let added = 0;
    const spy = vi.spyOn(pkg, 'addRelationship').mockImplementation((partUri, input) => {
      if (partUri === source.partUri && input.type === HYPERLINK_RELATIONSHIP) {
        added += 1;
        if (added === 2) throw new Error('injected rich cell relationship');
      }
      return originalAddRelationship(partUri, input);
    });
    try {
      expect(() => source.addTable([[
        { text: [{ runs: [
          { text: 'One', style: { hyperlink: { url: 'https://one.example' } } },
          { text: 'Two', style: { hyperlink: { url: 'https://two.example' } } },
        ] }] },
      ]])).toThrow('injected rich cell relationship');
      expect(packageSnapshot(pkg)).toEqual(beforeInvalid);
    } finally {
      spy.mockRestore();
    }
  });

  it('reads and whole-replaces table-cell rich text without partial plain edits', async () => {
    const { pkg, model } = emptyPresentationModel();
    const source = model.addSlide();
    model.addSlide();
    const table = source.addTable([[
      {
        text: [{
          align: 'center',
          runs: [
            { text: 'First', style: { bold: true } },
            { text: 'Soft', softBreakBefore: true },
          ],
        }, {
          runs: [{ text: 'Second', style: { italic: true } }],
        }],
        options: {
          fill: { kind: 'solid', color: { kind: 'srgb', value: 'DDEEFF' } },
          fit: 'shrink',
          margin: [1, 2, 3, 4],
          valign: 'middle',
        },
      },
      { text: 'Plain', options: { fill: { kind: 'none' } } },
    ]], { name: 'Editable rich table' });

    expect(table.rows[0]!.cells[0]!.text).toBe('First\nSoft\nSecond');
    expect(table.rows[0]!.cells[0]!.richText.map(({ runs }) =>
      runs.map(({ text }) => text))).toEqual([['First', 'Soft'], ['Second']]);
    expect(table.rows[0]!.cells[0]!.richText[0]!.runs[1]!.softBreakBefore).toBe(true);
    expect(table.rows[0]!.cells[0]!.richText[0]!.align).toBe('center');

    const snapshot = table.rows[0]!.cells[0]!.richText;
    const beforeNoOp = packageSnapshot(pkg);
    table.setCellRichText(0, 0, snapshot);
    expect(packageSnapshot(pkg)).toEqual(beforeNoOp);

    const replacement = [{
      align: 'right' as const,
      runs: [
        {
          text: 'URL',
          style: { hyperlink: { url: 'https://edit.example', tooltip: 'Edit' } },
        },
        {
          text: ' slide',
          style: { hyperlink: { slide: 2, tooltip: '' }, underline: false },
        },
      ],
    }, {
      runs: [{ text: 'Tail', style: { bold: true } }],
    }];
    const identity = table;
    table.setCellRichText(0, 0, replacement);
    expect(source.shapes.find(({ id }) => id === table.id)).toBe(identity);
    expect(table.rows[0]!.cells[0]!.text).toBe('URL slide\nTail');
    expect(table.rows[0]!.cells[0]!.richText[0]!.runs.map(
      ({ style }) => style?.hyperlink,
    )).toEqual([
      { url: 'https://edit.example', tooltip: 'Edit' },
      { slide: 2, tooltip: '' },
    ]);
    expect(table.rows[0]!.cells[0]!.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'DDEEFF' },
    });
    expect(table.rows[0]!.cells[0]!.margins).toEqual({
      top: 1,
      right: 2,
      bottom: 3,
      left: 4,
    });
    expect(table.rows[0]!.cells[0]!.textFit).toBe('shrink');
    expect(table.rows[0]!.cells[0]!.verticalAlignment).toBe('middle');
    expect(table.rows[0]!.cells[1]).toMatchObject({
      text: 'Plain',
      fill: { kind: 'none' },
    });

    replacement[0]!.runs[0]!.text = 'MUTATED';
    expect(table.rows[0]!.cells[0]!.text).toBe('URL slide\nTail');

    const beforeUnsafePlain = packageSnapshot(pkg);
    expect(() => table.setCellText(0, 0, 'Unsafe')).toThrow(ModelParseError);
    expect(packageSnapshot(pkg)).toEqual(beforeUnsafePlain);
    table.setCellText(0, 1, 'Safe');
    expect(table.rows[0]!.cells[1]!.text).toBe('Safe');

    const beforeInvalid = packageSnapshot(pkg);
    expect(() => table.setCellRichText(0, 0, [{ runs: [{
      text: 'Invalid',
      style: { hyperlink: { slide: 99 } },
    }] }])).toThrow();
    expect(() => table.setCellRichText(9, 0, [{ runs: [{ text: 'Missing' }] }]))
      .toThrow(RangeError);
    expect(packageSnapshot(pkg)).toEqual(beforeInvalid);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedTable = reopened.slides[0]!.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    )!;
    expect(reopenedTable.rows[0]!.cells[0]!.text).toBe('URL slide\nTail');
    expect(reopenedTable.rows[0]!.cells[0]!.richText[0]!.runs.map(
      ({ style }) => style?.hyperlink,
    )).toEqual([
      { url: 'https://edit.example', tooltip: 'Edit' },
      { slide: 2, tooltip: '' },
    ]);
  });

  it('edits rich table-cell links with reuse, COW, GC, and slide lifecycle safety', async () => {
    const { pkg, model } = emptyPresentationModel();
    const source = model.addSlide();
    const target = model.addSlide();
    model.addSlide();
    const table = source.addTable([[
      {
        text: [{ runs: [{ text: 'Shared one' }, { text: 'Shared two' }] }],
        options: { hyperlink: { url: 'https://shared.example' } },
      },
      { text: [{ runs: [{
        text: 'Unique',
        style: { hyperlink: { url: 'https://unique.example' } },
      }] }] },
      { text: [{ runs: [{
        text: 'Self',
        style: { hyperlink: { slide: 1, tooltip: '' } },
      }] }] },
      { text: [{ runs: [{
        text: 'Target',
        style: { hyperlink: { slide: 2 } },
      }] }] },
    ]], { name: 'Rich relationship lifecycle' });

    const cellClickIds = (cellIndex: number): readonly string[] => {
      const resolved = source.resolveShape(table.id);
      const row = resolved.xml.descendants(resolved.element, 'tr')[0]!;
      const cell = resolved.xml.descendants(row, 'tc')[cellIndex]!;
      return resolved.xml.descendants(cell, 'hlinkClick').map(
        (click) => resolved.xml.attribute(click, 'r:id')!.value,
      );
    };
    const editRun = (
      cellIndex: number,
      runIndex: number,
      hyperlink: Hyperlink | undefined,
    ): readonly RichTextParagraph[] => table.rows[0]!.cells[cellIndex]!.richText.map(
      (paragraph, paragraphIndex) => ({
        ...paragraph,
        runs: paragraph.runs.map((run, candidateIndex) => {
          if (paragraphIndex !== 0 || candidateIndex !== runIndex) return run;
          const { hyperlink: current, ...style } = run.style ?? {};
          void current;
          return {
            ...run,
            style: hyperlink === undefined ? style : { ...style, hyperlink },
          };
        }),
      }),
    );

    const [sharedId, secondSharedId] = cellClickIds(0);
    const uniqueId = cellClickIds(1)[0]!;
    expect(sharedId).toBe(secondSharedId);

    table.setCellRichText(0, 0, editRun(0, 0, {
      url: 'https://shared.example',
      tooltip: 'Updated',
    }));
    expect(cellClickIds(0)).toEqual([sharedId, sharedId]);
    expect(table.rows[0]!.cells[0]!.richText[0]!.runs[0]!.style?.hyperlink)
      .toEqual({ url: 'https://shared.example', tooltip: 'Updated' });

    table.setCellRichText(0, 1, editRun(1, 0, {
      url: 'https://unique-edited.example',
      tooltip: '',
    }));
    expect(cellClickIds(1)[0]).toBe(uniqueId);
    expect(source.relationships.find(({ id }) => id === uniqueId)?.target)
      .toBe('https://unique-edited.example');

    table.setCellRichText(0, 0, editRun(0, 0, {
      url: 'https://cloned.example',
    }));
    const [clonedId, retainedSharedId] = cellClickIds(0);
    expect(clonedId).not.toBe(sharedId);
    expect(retainedSharedId).toBe(sharedId);
    expect(source.relationships.find(({ id }) => id === clonedId)?.target)
      .toBe('https://cloned.example');
    expect(source.relationships.find(({ id }) => id === sharedId)?.target)
      .toBe('https://shared.example');

    table.setCellRichText(0, 0, editRun(0, 0, undefined));
    expect(source.relationships.some(({ id }) => id === clonedId)).toBe(false);
    expect(source.relationships.some(({ id }) => id === sharedId)).toBe(true);
    table.setCellRichText(0, 0, editRun(0, 1, undefined));
    expect(source.relationships.some(({ id }) => id === sharedId)).toBe(false);

    model.moveSlide(model.slides.indexOf(target), 0);
    expect(table.rows[0]!.cells[2]!.richText[0]!.runs[0]!.style?.hyperlink)
      .toEqual({ slide: 2, tooltip: '' });
    expect(table.rows[0]!.cells[3]!.richText[0]!.runs[0]!.style?.hyperlink)
      .toEqual({ slide: 1 });
    const duplicate = model.duplicateSlide(model.slides.indexOf(source));
    const duplicateTable = duplicate.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    )!;
    expect(duplicateTable.rows[0]!.cells[2]!.richText[0]!.runs[0]!.style?.hyperlink)
      .toEqual({ slide: model.slides.indexOf(duplicate) + 1, tooltip: '' });

    model.deleteSlide(model.slides.indexOf(target));
    expect(table.rows[0]!.cells[3]!.richText[0]!.runs[0]!.style?.hyperlink).toBeUndefined();
    expect(duplicateTable.rows[0]!.cells[3]!.richText[0]!.runs[0]!.style?.hyperlink)
      .toBeUndefined();

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSource = reopened.slides.find(({ partUri }) => partUri === source.partUri)!;
    const reopenedTable = reopenedSource.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    )!;
    expect(reopenedTable.rows[0]!.cells[1]!.richText[0]!.runs[0]!.style?.hyperlink)
      .toEqual({ url: 'https://unique-edited.example', tooltip: '' });
    expect(reopenedTable.rows[0]!.cells[2]!.richText[0]!.runs[0]!.style?.hyperlink)
      .toEqual({
        slide: reopened.slides.indexOf(reopenedSource) + 1,
        tooltip: '',
      });
  });

  it('rolls rich table-cell replacement back at every relationship mutation stage', () => {
    const { pkg, model } = emptyPresentationModel();
    const source = model.addSlide();
    model.addSlide();
    const table = source.addTable([[
      { text: [{ runs: [{
        text: 'Linked',
        style: { hyperlink: { url: 'https://original.example' } },
      }] }] },
      { text: [{ runs: [{ text: 'Plain' }] }] },
    ]]);
    const replaceFirst = (hyperlink: Hyperlink | undefined): readonly RichTextParagraph[] => {
      const [paragraph] = table.rows[0]!.cells[0]!.richText;
      const run = paragraph!.runs[0]!;
      const { hyperlink: current, ...style } = run.style ?? {};
      void current;
      return [{
        ...paragraph,
        runs: [{
          ...run,
          style: hyperlink === undefined ? style : { ...style, hyperlink },
        }],
      }];
    };
    const failureCases: readonly {
      readonly method: 'addRelationship' | 'updateRelationship' | 'removeRelationship' | 'setPart';
      readonly invoke: () => void;
    }[] = [
      {
        method: 'addRelationship',
        invoke: () => table.setCellRichText(0, 1, [{ runs: [{
          text: 'Added',
          style: { hyperlink: { url: 'https://added.example' } },
        }] }]),
      },
      {
        method: 'updateRelationship',
        invoke: () => table.setCellRichText(0, 0, replaceFirst({
          url: 'https://updated.example',
        })),
      },
      {
        method: 'removeRelationship',
        invoke: () => table.setCellRichText(0, 0, replaceFirst(undefined)),
      },
      {
        method: 'setPart',
        invoke: () => table.setCellRichText(0, 0, replaceFirst({
          url: 'https://original.example',
          tooltip: '',
        })),
      },
    ];

    for (const { method, invoke } of failureCases) {
      const before = packageSnapshot(pkg);
      const spy = vi.spyOn(pkg, method).mockImplementation(() => {
        throw new Error(`injected rich ${method}`);
      });
      try {
        expect(invoke).toThrow(`injected rich ${method}`);
        expect(packageSnapshot(pkg)).toEqual(before);
        expect(table.rows[0]!.cells[0]!.richText[0]!.runs[0]!.style?.hyperlink)
          .toEqual({ url: 'https://original.example' });
        expect(table.rows[0]!.cells[1]!.richText[0]!.runs[0]!.style?.hyperlink)
          .toBeUndefined();
      } finally {
        spy.mockRestore();
      }
    }

    const beforeOuter = packageSnapshot(pkg);
    expect(() => pkg.transaction(() => {
      table.setCellRichText(0, 1, [{ runs: [{
        text: 'Rollback',
        style: { hyperlink: { slide: 2, tooltip: '' } },
      }] }]);
      throw new Error('rollback rich table cell');
    })).toThrow('rollback rich table cell');
    expect(packageSnapshot(pkg)).toEqual(beforeOuter);
    expect(table.rows[0]!.cells[1]!.text).toBe('Plain');
  });

  it('edits table-cell hyperlinks with ID reuse, clone-on-write, and reference GC', async () => {
    const { pkg, model } = emptyPresentationModel();
    const source = model.addSlide();
    model.addSlide();
    const alternate = model.addSlide();
    const table = source.addTable([[
      { text: 'Shared one', options: { hyperlink: { url: 'https://shared.example' } } },
      { text: 'Shared two', options: { hyperlink: { url: 'https://second.example' } } },
      { text: 'Shared three', options: { hyperlink: { url: 'https://third.example' } } },
      { text: 'Target', options: { hyperlink: { slide: 2 } } },
      { text: 'Self', options: { hyperlink: { slide: 1, tooltip: '' } } },
      'Plain styled',
    ]], { name: 'Editable cell hyperlinks' });

    const prepared = source.resolveShape(table.id);
    const clicks = prepared.xml.descendants(prepared.element, 'hlinkClick');
    const firstId = prepared.xml.attribute(clicks[0]!, 'r:id')!.value;
    const replacedIds = clicks.slice(1, 3).map(
      (click) => prepared.xml.attribute(click, 'r:id')!.value,
    );
    for (const click of clicks.slice(1, 3)) {
      const id = prepared.xml.attribute(click, 'r:id')!.value;
      prepared.xml.replaceElement(
        click,
        prepared.xml.original(click).replace(`r:id="${id}"`, `r:id="${firstId}"`),
      );
    }
    const row = prepared.xml.descendants(prepared.element, 'tr')[0]!;
    const plainCell = prepared.xml.descendants(row, 'tc')[5]!;
    const plainProperties = prepared.xml.descendants(plainCell, 'rPr')[0]!;
    prepared.xml.replaceElement(
      plainProperties,
      prepared.xml.original(plainProperties).replace('<a:rPr ', '<a:rPr u="none" '),
    );
    pkg.transaction(() => {
      source.setXml(prepared.xml.serialize());
      for (const id of replacedIds) pkg.removeRelationship(source.partUri, id);
    });

    const directClickIds = (): readonly string[] => {
      const resolved = source.resolveShape(table.id);
      return resolved.xml.descendants(resolved.element, 'hlinkClick').map(
        (click) => resolved.xml.attribute(click, 'r:id')!.value,
      );
    };
    expect(table.rows[0]!.cells.slice(0, 3).map(({ hyperlink }) => hyperlink)).toEqual([
      { url: 'https://shared.example' },
      { url: 'https://shared.example' },
      { url: 'https://shared.example' },
    ]);

    const noOp = packageSnapshot(pkg);
    table.setCellHyperlink(0, 1, { url: 'https://shared.example' });
    expect(packageSnapshot(pkg)).toEqual(noOp);
    expect(source.shapes.find(({ id }) => id === table.id)).toBe(table);

    table.setCellHyperlink(0, 1, {
      url: 'https://shared.example',
      tooltip: '',
    });
    expect(directClickIds()[1]).toBe(firstId);
    expect(table.rows[0]!.cells[1]!.hyperlink).toEqual({
      url: 'https://shared.example',
      tooltip: '',
    });

    table.setCellHyperlink(0, 0, {
      url: 'https://edited.example',
      tooltip: 'Edited',
    });
    const clonedId = directClickIds()[0]!;
    expect(clonedId).not.toBe(firstId);
    expect(directClickIds().slice(1, 3)).toEqual([firstId, firstId]);
    expect(source.relationships.find(({ id }) => id === clonedId)?.target)
      .toBe('https://edited.example');
    expect(source.relationships.find(({ id }) => id === firstId)?.target)
      .toBe('https://shared.example');

    const uniqueInternalId = directClickIds()[3]!;
    table.setCellHyperlink(0, 3, { slide: 3, tooltip: '' });
    expect(directClickIds()[3]).toBe(uniqueInternalId);
    expect(source.relationships.find(({ id }) => id === uniqueInternalId)?.resolvedTarget)
      .toBe(alternate.partUri);

    table.setCellHyperlink(0, 5, { url: 'https://added.example' });
    let updatedXml = new TextDecoder().decode(pkg.requirePart(source.partUri).bytes);
    expect(updatedXml).toContain('<a:rPr u="none"');
    expect(updatedXml).not.toContain('<a:rPr u="none" u="sng"');
    table.setCellHyperlink(0, 5, undefined);
    updatedXml = new TextDecoder().decode(pkg.requirePart(source.partUri).bytes);
    expect(updatedXml).toContain('<a:rPr u="none"');
    expect(table.rows[0]!.cells[5]!.hyperlink).toBeUndefined();

    table.setCellHyperlink(0, 1, undefined);
    expect(source.relationships.some(({ id }) => id === firstId)).toBe(true);
    expect(table.rows[0]!.cells[2]!.hyperlink).toEqual({ url: 'https://shared.example' });
    table.setCellHyperlink(0, 2, undefined);
    expect(source.relationships.some(({ id }) => id === firstId)).toBe(false);

    table.setCellText(0, 0, 'Edited text');
    table.setCellBorders(0, 0, {
      kind: 'line',
      color: { kind: 'srgb', value: '112233' },
      width: 1,
    });
    table.setCellFill(0, 0, { kind: 'none' });
    table.setCellMargins(0, 0, { top: 2 });
    table.setCellTextDirection(0, 0, 'vert270');
    table.setCellTextFit(0, 0, 'shrink');
    table.setCellHorizontalAlignment(0, 0, 'center');
    table.setCellVerticalAlignment(0, 0, 'bottom');
    expect(table.rows[0]!.cells[0]!.hyperlink).toEqual({
      url: 'https://edited.example',
      tooltip: 'Edited',
    });

    model.moveSlide(model.slides.indexOf(alternate), 0);
    expect(table.rows[0]!.cells[3]!.hyperlink).toEqual({ slide: 1, tooltip: '' });
    expect(table.rows[0]!.cells[4]!.hyperlink).toEqual({ slide: 2, tooltip: '' });
    model.moveSlide(0, model.slides.length - 1);

    const duplicate = model.duplicateSlide(model.slides.indexOf(source));
    const duplicateTable = duplicate.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    )!;
    expect(duplicateTable.rows[0]!.cells[4]!.hyperlink).toEqual({
      slide: model.slides.indexOf(duplicate) + 1,
      tooltip: '',
    });
    model.deleteSlide(model.slides.indexOf(alternate));
    expect(table.rows[0]!.cells[3]!.hyperlink).toBeUndefined();
    expect(duplicateTable.rows[0]!.cells[3]!.hyperlink).toBeUndefined();

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSource = reopened.slides.find(({ partUri }) => partUri === source.partUri)!;
    const reopenedTable = reopenedSource.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    )!;
    expect(reopenedTable.rows[0]!.cells.map(({ hyperlink }) => hyperlink)).toEqual([
      { url: 'https://edited.example', tooltip: 'Edited' },
      undefined,
      undefined,
      undefined,
      { slide: reopened.slides.indexOf(reopenedSource) + 1, tooltip: '' },
      undefined,
    ]);
  });

  it('rejects unsafe table-cell hyperlink edits and rolls back every mutation stage', () => {
    const { pkg, model } = emptyPresentationModel();
    const source = model.addSlide();
    model.addSlide();
    const table = source.addTable([[
      { text: 'Linked', options: { hyperlink: { url: 'https://example.com' } } },
      'Plain',
    ]]);

    const invalidCalls = [
      () => table.setCellHyperlink(-1, 0, { url: 'https://invalid.example' }),
      () => table.setCellHyperlink(0, -1, { url: 'https://invalid.example' }),
      () => table.setCellHyperlink(1, 0, { url: 'https://invalid.example' }),
      () => table.setCellHyperlink(0, 2, { url: 'https://invalid.example' }),
      () => table.setCellHyperlink(0, 0, {} as never),
      () => table.setCellHyperlink(0, 0, { url: 'https://example.com', slide: 2 } as never),
      () => table.setCellHyperlink(0, 0, { slide: 99 }),
    ];
    for (const invoke of invalidCalls) {
      const before = packageSnapshot(pkg);
      expect(invoke).toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    const failureCases: readonly {
      readonly method: 'addRelationship' | 'updateRelationship' | 'removeRelationship' | 'setPart';
      readonly invoke: () => void;
    }[] = [
      {
        method: 'addRelationship',
        invoke: () => table.setCellHyperlink(0, 1, { url: 'https://added.example' }),
      },
      {
        method: 'updateRelationship',
        invoke: () => table.setCellHyperlink(0, 0, { url: 'https://updated.example' }),
      },
      {
        method: 'removeRelationship',
        invoke: () => table.setCellHyperlink(0, 0, undefined),
      },
      {
        method: 'setPart',
        invoke: () => table.setCellHyperlink(0, 0, {
          url: 'https://example.com',
          tooltip: '',
        }),
      },
    ];
    for (const { method, invoke } of failureCases) {
      const before = packageSnapshot(pkg);
      const spy = vi.spyOn(pkg, method).mockImplementation(() => {
        throw new Error(`injected ${method}`);
      });
      try {
        expect(invoke).toThrow(`injected ${method}`);
        expect(packageSnapshot(pkg)).toEqual(before);
        expect(table.rows[0]!.cells.map(({ hyperlink }) => hyperlink)).toEqual([
          { url: 'https://example.com' },
          undefined,
        ]);
      } finally {
        spy.mockRestore();
      }
    }

    const part = pkg.requirePart(source.partUri);
    const malformed = new TextDecoder().decode(part.bytes).replace(
      /<a:hlinkClick\b([^>]*)\/>/,
      '<a:hlinkClick$1 action="ppaction://unsupported"/>',
    );
    pkg.setPart(source.partUri, malformed, part.contentType);
    const beforeMalformed = packageSnapshot(pkg);
    expect(() => table.setCellHyperlink(0, 0, { url: 'https://replacement.example' }))
      .toThrow(ModelParseError);
    expect(packageSnapshot(pkg)).toEqual(beforeMalformed);

    pkg.setPart(source.partUri, part.bytes, part.contentType);
    const relationshipId = source.relationships.find(
      ({ type }) => type === HYPERLINK_RELATIONSHIP,
    )!.id;
    pkg.removeRelationship(source.partUri, relationshipId);
    const beforeDangling = packageSnapshot(pkg);
    expect(() => table.setCellHyperlink(0, 0, undefined)).toThrow(ModelParseError);
    expect(packageSnapshot(pkg)).toEqual(beforeDangling);
  });

  it('reopens table-cell hyperlinks in every presentation format', async () => {
    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const model = new PresentationModel(await OpcPackage.open(
        await modelFixture(profile.presentationContentType),
      ));
      const source = model.slides[0]!;
      const target = model.slides[1]!;
      const table = source.addTable([[
        { text: 'External', options: {
          hyperlink: { url: `https://example.com/${profile.format}`, tooltip: '' },
        } },
        { text: 'Internal', options: { hyperlink: { slide: 2 } } },
      ]], { name: 'Format links' });
      table.setCellHyperlink(0, 0, {
        url: `https://edited.example/${profile.format}`,
        tooltip: 'Edited',
      });
      table.setCellHyperlink(0, 1, { slide: 2, tooltip: '' });
      const richTable = source.addTable([[
        { text: [{ runs: [{ text: 'Rich external' }] }] },
        { text: [{ runs: [{ text: 'Rich internal' }] }] },
      ]], { name: 'Format rich links' });
      richTable.setCellRichText(0, 0, [{ runs: [{
        text: `Rich ${profile.format}`,
        style: { hyperlink: {
          url: `https://rich.example/${profile.format}`,
          tooltip: '',
        } },
      }] }]);
      richTable.setCellRichText(0, 1, [{ runs: [{
        text: 'Rich target',
        style: { hyperlink: { slide: 2, tooltip: 'Target' } },
      }] }]);
      const reopened = new PresentationModel(await OpcPackage.open(await model.opcPackage.write()));
      const reopenedTable = reopened.slides[0]!.shapes.find(
        (shape): shape is TableModel => shape instanceof TableModel
          && shape.name === 'Format links',
      )!;
      expect(reopened.format).toBe(profile.format);
      expect(reopenedTable.rows[0]!.cells.map(({ hyperlink }) => hyperlink)).toEqual([
        { url: `https://edited.example/${profile.format}`, tooltip: 'Edited' },
        {
          slide: reopened.slides.findIndex(({ partUri }) => partUri === target.partUri) + 1,
          tooltip: '',
        },
      ]);
      const reopenedRichTable = reopened.slides[0]!.shapes.find(
        (shape): shape is TableModel => shape instanceof TableModel
          && shape.name === 'Format rich links',
      )!;
      expect(reopenedRichTable.rows[0]!.cells.map(({ text }) => text)).toEqual([
        `Rich ${profile.format}`,
        'Rich target',
      ]);
      expect(reopenedRichTable.rows[0]!.cells.map(({ richText }) =>
        richText[0]!.runs[0]!.style?.hyperlink)).toEqual([
        { url: `https://rich.example/${profile.format}`, tooltip: '' },
        {
          slide: reopened.slides.findIndex(({ partUri }) => partUri === target.partUri) + 1,
          tooltip: 'Target',
        },
      ]);
    }
  });

  it('creates strict basic tables with stable identity, ordering, and atomic failure', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const part = pkg.requirePart(slide.partUri);
    pkg.setPart(
      slide.partUri,
      new TextDecoder().decode(part.bytes).replace(
        '</p:spTree>',
        '<p:extLst><p:ext uri="urn:test"><x:opaque xmlns:x="urn:test">KEEP</x:opaque></p:ext></p:extLst></p:spTree>',
      ),
      part.contentType,
    );

    const untouchedPartUri = '/ppt/custom/opaque1.bin';
    const untouchedBefore = pkg.requirePart(untouchedPartUri).bytes.slice();
    const sourceFillColor = { kind: 'srgb' as const, value: '#ff0000' };
    const sourceFill = {
      kind: 'solid' as const,
      color: sourceFillColor,
      transparency: 33.3334,
    };
    const sourceBorderColor = { kind: 'srgb' as const, value: '#c00000' };
    const sourceBorder = {
      kind: 'line' as const,
      color: sourceBorderColor,
      width: 2,
      style: 'solid' as const,
    };
    const sourceNamedMargin = { top: 4, left: 8 };
    const sourceTupleMargin: [number, number, number, number] = [1, 2, 3, 4];
    const typedSourceMargin: TextBoxMarginInput = sourceTupleMargin;
    const typedSourceBorder: TableCellBorderInput = sourceBorder;
    const sourceCellOptions: AddTableCellOptions = {
      align: 'center',
      border: typedSourceBorder,
      fill: sourceFill,
      margin: sourceNamedMargin,
      textDirection: 'vert',
      valign: 'middle',
    };
    const objectCell = { text: 'A & <1>', options: sourceCellOptions };
    const tableRows: readonly (readonly AddTableCellInput[])[] = [
      [
        objectCell,
        { text: '', options: {
          border: [
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
          ],
          fill: { kind: 'none' },
          margin: typedSourceMargin,
          textDirection: 'vert270',
          valign: 'top',
          align: 'left',
        } },
        {
          text: 'Empty margin',
          options: {
            align: 'right',
            margin: {},
            textDirection: 'wordArtVert',
            valign: 'bottom',
          },
        },
      ],
      [{
        text: 'A2',
        options: {
          border: {
            top: {
              kind: 'line',
              color: { kind: 'scheme', value: 'accent3' },
              width: 1584,
            },
            left: { kind: 'none' },
          },
          fill: {
            kind: 'solid',
            color: { kind: 'scheme', value: 'accent2' },
            transparency: 25,
          },
          margin: 6,
          textDirection: 'horz',
          valign: 'bottom',
          align: 'justify',
        },
      }, {
        text: 'B2',
        options: {
          align: undefined,
          border: { kind: 'none' },
          margin: 0,
        } as unknown as AddTableCellOptions,
      }, 'String'],
    ];
    const first = slide.addText('Before table', { name: 'Before' });
    const tableOptions: AddTableOptions = {
      align: 'center',
      name: 'Table "A"',
      x: inches(1),
      y: inches(1.5),
      width: inches(4),
      height: inches(2),
      columnWidths: [inches(1), inches(1), inches(2)],
      rowHeights: [inches(0.75), inches(1.25)],
    };
    const table = slide.addTable(tableRows, tableOptions);
    const last = slide.addText('After table', { name: 'After' });
    const directAlignments = (partBytes: Uint8Array): readonly (string | undefined)[] => {
      const xml = new TextDecoder().decode(partBytes);
      const nameOffset = xml.indexOf('name="Table &quot;A&quot;"');
      const frameStart = xml.lastIndexOf('<p:graphicFrame', nameOffset);
      const frameEnd = xml.indexOf('</p:graphicFrame>', nameOffset);
      expect(nameOffset).toBeGreaterThanOrEqual(0);
      expect(frameStart).toBeGreaterThanOrEqual(0);
      expect(frameEnd).toBeGreaterThan(nameOffset);
      const frame = xml.slice(frameStart, frameEnd + '</p:graphicFrame>'.length);
      return [...frame.matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g)]
        .map((match) => match[0].match(
          /<a:pPr[^>]*\salgn="([^"]+)"/,
        )?.[1]);
    };
    const directDirections = (partBytes: Uint8Array): readonly (string | undefined)[] => {
      const xml = new TextDecoder().decode(partBytes);
      const nameOffset = xml.indexOf('name="Table &quot;A&quot;"');
      const frameStart = xml.lastIndexOf('<p:graphicFrame', nameOffset);
      const frameEnd = xml.indexOf('</p:graphicFrame>', nameOffset);
      const frame = xml.slice(frameStart, frameEnd + '</p:graphicFrame>'.length);
      return [...frame.matchAll(/<a:tc(?:\s[^>]*)?>[\s\S]*?<\/a:tc>/g)]
        .map((match) => match[0].match(
          /<a:tcPr[^>]*\svert="([^"]+)"/,
        )?.[1]);
    };

    expect([first.id, table.id, last.id]).toEqual([2, 3, 4]);
    expect(table).toBeInstanceOf(TableModel);
    expect(table.name).toBe('Table "A"');
    expect(table.rows.map(({ cells }) => cells.map(({ text }) => text))).toEqual([
      ['A & <1>', '', 'Empty margin'],
      ['A2', 'B2', 'String'],
    ]);
    expect(table.rows.map(({ cells }) => cells.map(({ fill }) => fill))).toEqual([
      [
        {
          kind: 'solid',
          color: { kind: 'srgb', value: 'FF0000' },
          transparency: 33.333,
        },
        { kind: 'none' },
        undefined,
      ],
      [
        {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 25,
        },
        undefined,
        undefined,
      ],
    ]);
    expect(table.rows.map(({ cells }) => cells.map(({ margins }) => margins))).toEqual([
      [
        { top: 4, right: 7.2, bottom: 3.6, left: 8 },
        { top: 1, right: 2, bottom: 3, left: 4 },
        { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 },
      ],
      [
        { top: 6, right: 6, bottom: 6, left: 6 },
        { top: 0, right: 0, bottom: 0, left: 0 },
        { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 },
      ],
    ]);
    expect(table.rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment))).toEqual([
      ['middle', 'top', 'bottom'],
      ['bottom', undefined, undefined],
    ]);
    const expectedDirections = [
      'vert',
      'vert270',
      'wordArtVert',
      undefined,
      undefined,
      undefined,
    ];
    expect(table.rows.flatMap(({ cells }) =>
      cells.map(({ textDirection }) => textDirection))).toEqual(expectedDirections);
    expect(directDirections(pkg.requirePart(slide.partUri).bytes)).toEqual(
      expectedDirections,
    );
    const expectedAlignments = ['ctr', 'l', 'r', 'just', 'ctr', 'ctr'];
    expect(directAlignments(pkg.requirePart(slide.partUri).bytes)).toEqual(
      expectedAlignments,
    );
    const sourceIndex = model.slides.findIndex(({ partUri }) => partUri === slide.partUri);
    const duplicatedSlide = model.duplicateSlide(sourceIndex);
    const duplicatedTable = duplicatedSlide.shapes.find(
      (shape) => shape.name === table.name,
    );
    expect(duplicatedTable).toBeInstanceOf(TableModel);
    expect(directAlignments(pkg.requirePart(duplicatedSlide.partUri).bytes)).toEqual(
      expectedAlignments,
    );
    expect(directDirections(pkg.requirePart(duplicatedSlide.partUri).bytes)).toEqual(
      expectedDirections,
    );
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
    expect(table.rows.map(({ cells }) => cells.map(({ borders }) => borders))).toEqual([
      [
        {
          top: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
          right: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
          bottom: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
          left: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
        },
        {
          top: borderLine({ kind: 'scheme', value: 'accent1' }, 1.5, 'dash'),
          right: { kind: 'none' },
          bottom: borderLine({ kind: 'srgb', value: '00FF00' }, 0),
          left: { kind: 'none' },
        },
        noBorders(),
      ],
      [
        {
          top: borderLine({ kind: 'scheme', value: 'accent3' }, 1584),
          right: { kind: 'none' },
          bottom: { kind: 'none' },
          left: { kind: 'none' },
        },
        noBorders(),
        noBorders(),
      ],
    ]);
    objectCell.text = 'MUTATED';
    sourceFillColor.value = '000000';
    sourceFill.transparency = 1;
    sourceBorderColor.value = '000000';
    sourceBorder.width = 9;
    sourceNamedMargin.top = 99;
    sourceNamedMargin.left = 99;
    sourceTupleMargin[0] = 99;
    sourceTupleMargin[3] = 99;
    expect(table.rows[0]!.cells[0]!.text).toBe('A & <1>');
    expect(table.rows[0]!.cells[0]!.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF0000' },
      transparency: 33.333,
    });
    expect(table.rows[0]!.cells[0]!.margins).toEqual({
      top: 4,
      right: 7.2,
      bottom: 3.6,
      left: 8,
    });
    expect(table.rows[0]!.cells[1]!.margins).toEqual({
      top: 1,
      right: 2,
      bottom: 3,
      left: 4,
    });
    expect(table.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual([
      'middle',
      'top',
      'bottom',
    ]);
    expect(table.rows[0]!.cells[0]!.borders).toEqual({
      top: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
      right: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
      bottom: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
      left: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
    });
    expect(table.transform).toMatchObject({
      x: inches(1),
      y: inches(1.5),
      width: inches(4),
      height: inches(2),
    });
    expect(table.columnWidths).toEqual([inches(1), inches(1), inches(2)]);
    expect(table.rowHeights).toEqual([inches(0.75), inches(1.25)]);
    expect(slide.shapes.find(({ id }) => id === table.id)).toBe(table);

    table.setCellMargins(0, 0, { top: 2 });
    expect(table.rows[0]!.cells[0]!.margins).toEqual({ top: 2 });
    table.setCellMargins(0, 1, undefined);
    expect(table.rows[0]!.cells[1]!.margins).toBeUndefined();
    table.setCellVerticalAlignment(0, 0, 'bottom');
    table.setCellVerticalAlignment(0, 1, undefined);
    expect(table.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual([
      'bottom',
      undefined,
      'bottom',
    ]);
    expect(table.rows[0]!.cells[0]!.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF0000' },
      transparency: 33.333,
    });
    expect(table.rows[0]!.cells[0]!.borders).toEqual({
      top: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
      right: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
      bottom: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
      left: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
    });

    table.setCellBorders(0, 1, {
      right: borderLine({ kind: 'scheme', value: 'accent6' }, 2.25, 'dash'),
    });
    expect(table.rows[0]!.cells[1]!.borders).toEqual({
      right: borderLine({ kind: 'scheme', value: 'accent6' }, 2.25, 'dash'),
    });
    expect(table.rows[0]!.cells[1]!.fill).toEqual({ kind: 'none' });
    table.setCellBorders(0, 1, undefined);
    expect(table.rows[0]!.cells[1]!.borders).toBeUndefined();
    expect(table.rows[0]!.cells[1]!.fill).toEqual({ kind: 'none' });
    table.setCellFill(0, 1, {
      kind: 'solid',
      color: { kind: 'srgb', value: '00FF00' },
    });
    expect(table.rows[0]!.cells[1]!.fill).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '00FF00' },
    });
    table.setCellFill(0, 1, undefined);
    expect(table.rows[0]!.cells[1]!.fill).toBeUndefined();
    table.setCellText(0, 0, 'Edited');
    table.setTransform({ x: inches(2) });
    expect(slide.shapes.find(({ id }) => id === table.id)).toBe(table);
    expect(table.rows[0]!.cells[0]!.text).toBe('Edited');
    expect(table.transform.x).toBe(inches(2));
    expect(table.columnWidths).toEqual([inches(1), inches(1), inches(2)]);
    expect(table.rowHeights).toEqual([inches(0.75), inches(1.25)]);
    expect(directAlignments(pkg.requirePart(slide.partUri).bytes)).toEqual(
      expectedAlignments,
    );
    expect(directAlignments(pkg.requirePart(duplicatedSlide.partUri).bytes)).toEqual(
      expectedAlignments,
    );
    expect(directDirections(pkg.requirePart(slide.partUri).bytes)).toEqual(
      expectedDirections,
    );
    expect(directDirections(pkg.requirePart(duplicatedSlide.partUri).bytes)).toEqual(
      expectedDirections,
    );

    const updated = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(updated.indexOf('name="Before"')).toBeLessThan(updated.indexOf('name="Table &quot;A&quot;"'));
    expect(updated.indexOf('name="Table &quot;A&quot;"')).toBeLessThan(updated.indexOf('name="After"'));
    expect(updated.indexOf('name="After"')).toBeLessThan(updated.indexOf('<p:extLst>'));
    expect(updated).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');

    const unequal = slide.addTable(
      [['A', 'B', 'C'], ['D', 'E', 'F']],
      {
        name: 'Unequal columns',
        columnWidths: [inches(1), inches(2), inches(3)],
        rowHeights: [inches(0.5), inches(1.5)],
      },
    );
    expect(unequal.id).toBe(5);
    expect(unequal.transform.width).toBe(inches(6));
    expect(unequal.transform.height).toBe(inches(2));
    const unequalSlideXml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    const unequalNameOffset = unequalSlideXml.indexOf('name="Unequal columns"');
    const unequalFrameStart = unequalSlideXml.lastIndexOf('<p:graphicFrame', unequalNameOffset);
    const unequalFrameEnd = unequalSlideXml.indexOf('</p:graphicFrame>', unequalNameOffset);
    const unequalFrame = unequalSlideXml.slice(
      unequalFrameStart,
      unequalFrameEnd + '</p:graphicFrame>'.length,
    );
    expect(unequalFrameStart).toBeGreaterThanOrEqual(0);
    expect(unequalFrameEnd).toBeGreaterThan(unequalNameOffset);
    expect([...unequalFrame.matchAll(/<a:gridCol w="(\d+)"\/>/g)]
      .map((match) => Number(match[1]))).toEqual([
      inches(1),
      inches(2),
      inches(3),
    ]);
    expect(unequalFrame).toContain('<a:ext cx="5486400" cy="1828800"/>');
    expect([...unequalFrame.matchAll(/<a:tr h="(\d+)">/g)]
      .map((match) => Number(match[1]))).toEqual([
      inches(0.5),
      inches(1.5),
    ]);
    expect(unequalFrame.match(/<a:tc>/g)).toHaveLength(6);

    const beforeInvalid = pkg.requirePart(slide.partUri).bytes;
    const invalidJournal = [...pkg.mutations];
    const shapeCount = slide.shapes.length;
    const accessorCell = {};
    let cellAccessorCalls = 0;
    Object.defineProperty(accessorCell, 'text', {
      get() {
        cellAccessorCalls += 1;
        return 'Accessor';
      },
      enumerable: true,
      configurable: true,
    });
    const accessorBorderOptions = {};
    let borderAccessorCalls = 0;
    Object.defineProperty(accessorBorderOptions, 'border', {
      get() {
        borderAccessorCalls += 1;
        return { kind: 'none' };
      },
      enumerable: true,
      configurable: true,
    });
    const invalidRows = [
      null,
      [],
      ['A'],
      [['A'], ['B', 'C']],
      [[1]],
      [[{ text: 'A', unknown: true }]],
      [[{ text: 'A', options: { fill: {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
        transparency: 101,
      } } }]],
      [[{ text: 'A', options: accessorBorderOptions }]],
      [[{ text: 'A', options: { border: { kind: 'unknown' } } }]],
      [[{ text: 'A', options: { border: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FFF' },
        width: 1,
      } } }]],
      [[{ text: 'A', options: { border: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        width: -1,
      } } }]],
      [[{ text: 'A', options: { border: [
        { kind: 'none' },
        undefined,
        undefined,
      ] } }]],
      [[accessorCell]],
    ];
    for (const rows of invalidRows) {
      expect(() => slide.addTable(rows as never)).toThrow();
    }
    for (const fit of [null, false, 0, '', 'Shrink', ' shrink', 'auto', [], {}]) {
      expect(() => slide.addTable([[
        { text: 'Invalid fit', options: { fit: fit as never } },
      ]])).toThrow(TypeError);
    }
    expect(cellAccessorCalls).toBe(0);
    expect(borderAccessorCalls).toBe(0);
    const invalidOptions: unknown[] = [
      null,
      [],
      { unknown: true },
      { name: 1 },
      { x: Number.NaN },
      { width: 0 },
      { height: Number.POSITIVE_INFINITY },
      { columnWidths: [1, 2] },
      { width: 4, columnWidths: [1] },
      { rowHeights: [1, 2] },
      { height: 4, rowHeights: [1] },
    ];
    const accessorTableAlign = {};
    let tableAlignAccessorCalls = 0;
    Object.defineProperty(accessorTableAlign, 'align', {
      get() {
        tableAlignAccessorCalls += 1;
        return 'center';
      },
      enumerable: true,
      configurable: true,
    });
    invalidOptions.push(accessorTableAlign);
    for (const align of [
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
    ]) {
      invalidOptions.push({ align });
    }
    const accessorColumnWidths = [1];
    let columnWidthAccessorCalls = 0;
    Object.defineProperty(accessorColumnWidths, '0', {
      get() {
        columnWidthAccessorCalls += 1;
        return 1;
      },
      enumerable: true,
      configurable: true,
    });
    invalidOptions.push({ columnWidths: accessorColumnWidths });
    const accessorRowHeights = [1];
    let rowHeightAccessorCalls = 0;
    Object.defineProperty(accessorRowHeights, '0', {
      get() {
        rowHeightAccessorCalls += 1;
        return 1;
      },
      enumerable: true,
      configurable: true,
    });
    invalidOptions.push({ rowHeights: accessorRowHeights });
    for (const options of invalidOptions) {
      expect(() => slide.addTable([['A']], options as never)).toThrow();
    }
    expect(columnWidthAccessorCalls).toBe(0);
    expect(rowHeightAccessorCalls).toBe(0);
    expect(tableAlignAccessorCalls).toBe(0);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);
    expect(slide.shapes).toHaveLength(shapeCount);
    expect(slide.shapes[1]).toBe(table);

    let rolledBack: TableModel | undefined;
    const beforeRollback = pkg.requirePart(slide.partUri).bytes;
    const rollbackJournal = [...pkg.mutations];
    expect(() =>
      pkg.transaction(() => {
        rolledBack = slide.addTable([[{
          text: 'rollback',
          options: {
            border: {
              kind: 'line',
              color: { kind: 'scheme', value: 'accent1' },
              width: 2,
              style: 'dash',
            },
            fill: {
              kind: 'solid',
              color: { kind: 'scheme', value: 'accent1' },
              transparency: 25,
            },
            margin: { right: -2 },
            textDirection: 'wordArtVert',
            valign: 'top',
            align: 'right',
          },
        }]]);
        throw new Error('restore table');
      }),
    ).toThrow('restore table');
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(() => rolledBack!.rows).toThrow(ModelParseError);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri);
    const reopenedTable = reopenedSlide?.shapes.find(({ id }) => id === table.id);
    const reopenedDuplicatedSlide = reopened.slides.find(
      ({ partUri }) => partUri === duplicatedSlide.partUri,
    );
    expect(reopenedTable).toBeInstanceOf(TableModel);
    expect(reopenedSlide!.shapes.find(({ id }) => id === table.id)).toBe(reopenedTable);
    expect((reopenedTable as TableModel).rows.map(({ cells }) =>
      cells.map(({ text }) => text))).toEqual([
      ['Edited', '', 'Empty margin'],
      ['A2', 'B2', 'String'],
    ]);
    expect((reopenedTable as TableModel).rows.map(({ cells }) =>
      cells.map(({ margins }) => margins))).toEqual([
      [
        { top: 2 },
        undefined,
        { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 },
      ],
      [
        { top: 6, right: 6, bottom: 6, left: 6 },
        { top: 0, right: 0, bottom: 0, left: 0 },
        { top: 3.6, right: 7.2, bottom: 3.6, left: 7.2 },
      ],
    ]);
    expect((reopenedTable as TableModel).rows.map(({ cells }) =>
      cells.map(({ verticalAlignment }) => verticalAlignment))).toEqual([
      ['bottom', undefined, 'bottom'],
      ['bottom', undefined, undefined],
    ]);
    expect((reopenedTable as TableModel).rows.flatMap(({ cells }) =>
      cells.map(({ textDirection }) => textDirection))).toEqual(expectedDirections);
    expect((reopenedTable as TableModel).rows.map(({ cells }) =>
      cells.map(({ fill }) => fill))).toEqual([
      [
        {
          kind: 'solid',
          color: { kind: 'srgb', value: 'FF0000' },
          transparency: 33.333,
        },
        undefined,
        undefined,
      ],
      [
        {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 25,
        },
        undefined,
        undefined,
      ],
    ]);
    expect((reopenedTable as TableModel).rows.map(({ cells }) =>
      cells.map(({ borders }) => borders))).toEqual([
      [
        {
          top: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
          right: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
          bottom: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
          left: borderLine({ kind: 'srgb', value: 'C00000' }, 2, 'solid'),
        },
        undefined,
        noBorders(),
      ],
      [
        {
          top: borderLine({ kind: 'scheme', value: 'accent3' }, 1584),
          right: { kind: 'none' },
          bottom: { kind: 'none' },
          left: { kind: 'none' },
        },
        noBorders(),
        noBorders(),
      ],
    ]);
    expect((reopenedTable as TableModel).columnWidths).toEqual([
      inches(1),
      inches(1),
      inches(2),
    ]);
    expect((reopenedTable as TableModel).rowHeights).toEqual([
      inches(0.75),
      inches(1.25),
    ]);
    expect(reopenedTable!.transform).toMatchObject({
      x: inches(2),
      y: inches(1.5),
      width: inches(4),
      height: inches(2),
    });
    expect(directAlignments(
      reopened.opcPackage.requirePart(reopenedSlide!.partUri).bytes,
    )).toEqual(expectedAlignments);
    expect(directAlignments(
      reopened.opcPackage.requirePart(reopenedDuplicatedSlide!.partUri).bytes,
    )).toEqual(expectedAlignments);
    expect(directDirections(
      reopened.opcPackage.requirePart(reopenedSlide!.partUri).bytes,
    )).toEqual(expectedDirections);
    expect(directDirections(
      reopened.opcPackage.requirePart(reopenedDuplicatedSlide!.partUri).bytes,
    )).toEqual(expectedDirections);
    expect(pkg.requirePart(untouchedPartUri).bytes).toEqual(untouchedBefore);
    expect(reopened.opcPackage.requirePart(untouchedPartUri).bytes).toEqual(untouchedBefore);

    const missingTree = model.addSlide();
    const missingPart = pkg.requirePart(missingTree.partUri);
    pkg.setPart(
      missingTree.partUri,
      '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld/></p:sld>',
      missingPart.contentType,
    );
    const beforeMissing = pkg.requirePart(missingTree.partUri).bytes;
    const missingJournal = [...pkg.mutations];
    expect(() => missingTree.addTable([['A'], ['B', 'C']])).toThrow(/same number/);
    expect(() => missingTree.addTable([['A']])).toThrow(/exactly one direct shape tree/);
    expect(pkg.requirePart(missingTree.partUri).bytes).toEqual(beforeMissing);
    expect(pkg.mutations).toEqual(missingJournal);

    const repeatedTree = model.addSlide();
    const repeatedTreePart = pkg.requirePart(repeatedTree.partUri);
    pkg.setPart(
      repeatedTree.partUri,
      '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/><p:spTree/></p:cSld></p:sld>',
      repeatedTreePart.contentType,
    );
    expect(() => repeatedTree.addTable([['A']])).toThrow(/exactly one direct shape tree/);

    const repeatedExtension = model.addSlide();
    const repeatedExtensionPart = pkg.requirePart(repeatedExtension.partUri);
    pkg.setPart(
      repeatedExtension.partUri,
      new TextDecoder().decode(repeatedExtensionPart.bytes).replace(
        '</p:spTree>',
        '<p:extLst/><p:extLst/></p:spTree>',
      ),
      repeatedExtensionPart.contentType,
    );
    expect(() => repeatedExtension.addTable([['A']])).toThrow(/repeated extension lists/);
  });

  it('creates logical table spans with physical hyperlink ownership and reopen', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const table = slide.addTable([
      [{
        text: [{ runs: [
          { text: 'Default' },
          {
            text: ' run',
            style: { hyperlink: { url: 'https://run.example' } },
          },
        ] }],
        options: {
          colspan: 2,
          rowspan: 2,
          hyperlink: { url: 'https://default.example' },
        },
      }, 'Right'],
      ['Bottom right'],
    ], { name: 'Merged creation' });

    expect(table.rows.map(({ cells }) => cells.map(({ text }) => text))).toEqual([
      ['Default run', '', 'Right'],
      ['', '', 'Bottom right'],
    ]);
    const xml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(xml).toContain('<a:tc rowSpan="2" gridSpan="2"><a:txBody>');
    expect(xml).toContain('<a:tc rowSpan="2" hMerge="1"><a:tcPr/></a:tc>');
    expect(xml).toContain('<a:tc gridSpan="2" vMerge="1"><a:tcPr/></a:tc>');
    expect(xml).toContain('<a:tc vMerge="1" hMerge="1"><a:tcPr/></a:tc>');
    expect(xml.match(/<a:tc\b/g)).toHaveLength(6);
    expect(xml.match(/<a:hlinkClick\b/g)).toHaveLength(2);
    expect(slide.relationships.filter(({ type }) => type === HYPERLINK_RELATIONSHIP))
      .toHaveLength(2);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedTable = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!
      .shapes.find((shape): shape is TableModel =>
        shape instanceof TableModel && shape.name === 'Merged creation')!;
    expect(reopenedTable.rows.map(({ cells }) => cells.map(({ text }) => text))).toEqual([
      ['Default run', '', 'Right'],
      ['', '', 'Bottom right'],
    ]);
    expect(reopened.slides.find(({ partUri }) => partUri === slide.partUri)!
      .relationships.filter(({ type }) => type === HYPERLINK_RELATIONSHIP))
      .toHaveLength(2);

    const beforeInvalid = packageSnapshot(pkg);
    expect(() => slide.addTable([[
      { text: 'Invalid', options: { colspan: 2, rowspan: 2 } },
    ]])).toThrow(/span is out of range/);
    expect(packageSnapshot(pkg)).toEqual(beforeInvalid);
  });

  it('preserves table-cell text style defaults through edits, duplication, rollback, and reopen', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const untouchedPartUri = '/ppt/custom/opaque1.bin';
    const untouchedBefore = pkg.requirePart(untouchedPartUri).bytes.slice();
    const table = slide.addTable([[
      'Plain',
      { text: 'False', options: { bold: false, spacing: { before: 3 } } },
      {
        text: [{
          spacing: { after: 9, line: false },
          runs: [
            { text: 'Inherited' },
            { text: ' local', style: { fontSize: 10, bold: false } },
          ],
        }, { spacing: false, runs: [] }],
        options: {
          fontFamily: 'Courier New',
          color: { kind: 'srgb', value: '00AA00' },
        },
      },
      { text: [{ runs: [] }] },
      { text: 'Linked', options: {
        hyperlink: { url: 'https://defaults.example' },
      } },
    ]], {
      name: 'Table-cell text style defaults',
      fontFamily: 'Aptos',
      fontSize: 18.25,
      bold: true,
      color: { kind: 'scheme', value: 'accent1' },
      spacing: {
        before: 6,
        after: 8,
        line: { kind: 'multiple', factor: 1.5 },
      },
    });

    expect(table.rows[0]!.cells[0]!.richText[0]!.runs[0]!.style).toMatchObject({
      fontFamily: 'Aptos',
      fontSize: 18.25,
      bold: true,
      color: { kind: 'scheme', value: 'accent1' },
    });
    expect(table.rows[0]!.cells[1]!.richText[0]!.runs[0]!.style?.bold).toBe(false);
    expect(table.rows[0]!.cells[1]!.richText[0]!.spacing).toEqual({
      before: 3,
      after: 8,
      line: { kind: 'multiple', factor: 1.5 },
    });
    expect(table.rows[0]!.cells[2]!.richText).toMatchObject([{
      spacing: { before: 6, after: 9 },
      runs: [
        {
          text: 'Inherited',
          style: {
            fontFamily: 'Courier New',
            fontSize: 18.25,
            bold: true,
            color: { kind: 'srgb', value: '00AA00' },
          },
        },
        {
          text: ' local',
          style: {
            fontFamily: 'Courier New',
            fontSize: 10,
            bold: false,
            color: { kind: 'srgb', value: '00AA00' },
          },
        },
      ],
    }, { runs: [] }]);
    expect(table.rows[0]!.cells[3]!.richText).toMatchObject([{
      runs: [],
      spacing: {
        before: 6,
        after: 8,
        line: { kind: 'multiple', factor: 1.5 },
      },
    }]);
    expect(table.rows[0]!.cells[4]!.richText[0]!.runs[0]!.style).toMatchObject({
      fontFamily: 'Aptos',
      fontSize: 18.25,
      bold: true,
      color: { kind: 'scheme', value: 'accent1' },
      hyperlink: { url: 'https://defaults.example' },
    });

    const createdState = table.rows[0]!.cells.map(({ richText }) => richText);
    const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
    const duplicateTable = duplicate.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel
        && shape.name === table.name,
    )!;
    expect(duplicateTable.rows[0]!.cells.map(({ richText }) => richText)).toEqual(createdState);

    const noOpBefore = pkg.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...pkg.mutations];
    table.setCellRichText(0, 2, table.rows[0]!.cells[2]!.richText);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(noOpBefore);
    expect(pkg.mutations).toEqual(noOpJournal);

    const plainStyle = table.rows[0]!.cells[0]!.richText[0]!.runs[0]!.style;
    table.setCellText(0, 0, 'Edited plain');
    expect(table.rows[0]!.cells[0]!.richText[0]!.runs[0]).toEqual({
      text: 'Edited plain',
      style: plainStyle,
    });
    table.setCellRichText(0, 1, [{ runs: [{ text: 'Replacement' }] }]);
    const replacement = table.rows[0]!.cells[1]!.richText[0]!;
    expect(replacement.spacing).toBeUndefined();
    expect(replacement.runs[0]!.style).toMatchObject({
      fontFamily: '+mn-lt',
      color: { kind: 'scheme', value: 'tx1' },
    });
    expect(replacement.runs[0]!.style?.fontSize).toBeUndefined();
    expect(replacement.runs[0]!.style?.bold).toBeUndefined();
    expect(duplicateTable.rows[0]!.cells.map(({ richText }) => richText)).toEqual(createdState);

    const editedState = table.rows[0]!.cells.map(({ richText }) => richText);
    const beforeRollback = packageSnapshot(pkg);
    let rolledBack: TableModel | undefined;
    expect(() => pkg.transaction(() => {
      table.setCellRichText(0, 0, [{ runs: [{
        text: 'Temporary',
        style: { fontFamily: 'Arial', bold: false },
      }] }]);
      rolledBack = slide.addTable([['Temporary table']], {
        fontFamily: 'Arial',
        fontSize: 12,
        bold: false,
        color: { kind: 'srgb', value: 'FF0000' },
        spacing: { after: 4 },
      });
      throw new Error('restore table-cell text style defaults');
    })).toThrow('restore table-cell text style defaults');
    expect(packageSnapshot(pkg)).toEqual(beforeRollback);
    expect(table.rows[0]!.cells.map(({ richText }) => richText)).toEqual(editedState);
    expect(() => rolledBack!.rows).toThrow(ModelParseError);
    expect(pkg.requirePart(untouchedPartUri).bytes).toEqual(untouchedBefore);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSource = reopened.slides.find(({ partUri }) => partUri === slide.partUri)!;
    const reopenedDuplicate = reopened.slides.find(({ partUri }) => partUri === duplicate.partUri)!;
    const reopenedTable = reopenedSource.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel
        && shape.name === table.name,
    )!;
    const reopenedDuplicateTable = reopenedDuplicate.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel
        && shape.name === table.name,
    )!;
    expect(reopenedTable.rows[0]!.cells.map(({ richText }) => richText)).toEqual(editedState);
    expect(reopenedDuplicateTable.rows[0]!.cells.map(({ richText }) => richText))
      .toEqual(createdState);
    expect(reopened.opcPackage.requirePart(untouchedPartUri).bytes).toEqual(untouchedBefore);
  });

  it('reopens table-cell text style defaults in every presentation format', async () => {
    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const model = new PresentationModel(await OpcPackage.open(
        await modelFixture(profile.presentationContentType),
      ));
      const table = model.addSlide().addTable([[
        'Inherited',
        { text: 'False', options: { bold: false, spacing: { after: 3 } } },
      ]], {
        name: `Text defaults ${profile.format}`,
        fontFamily: 'Aptos',
        fontSize: 18.25,
        bold: true,
        color: { kind: 'scheme', value: 'accent1' },
        spacing: { before: 6, line: { kind: 'multiple', factor: 1.5 } },
      });
      table.setCellText(0, 0, `Edited ${profile.format}`);

      const reopened = new PresentationModel(await OpcPackage.open(
        await model.opcPackage.write(),
      ));
      const reopenedTable = reopened.slides.at(-1)!.shapes[0] as TableModel;
      expect(reopened.format).toBe(profile.format);
      expect(reopenedTable.rows[0]!.cells[0]!.richText[0]!.runs[0]).toMatchObject({
        text: `Edited ${profile.format}`,
        style: {
          fontFamily: 'Aptos',
          fontSize: 18.25,
          bold: true,
          color: { kind: 'scheme', value: 'accent1' },
        },
      });
      expect(reopenedTable.rows[0]!.cells[0]!.richText[0]!.spacing).toEqual({
        before: 6,
        line: { kind: 'multiple', factor: 1.5 },
      });
      expect(reopenedTable.rows[0]!.cells[1]!.richText[0]!.runs[0]!.style?.bold)
        .toBe(false);
      expect(reopenedTable.rows[0]!.cells[1]!.richText[0]!.spacing).toEqual({
        before: 6,
        after: 3,
        line: { kind: 'multiple', factor: 1.5 },
      });
    }
  });

  it('materializes table valign through edit, duplicate, write, and reopen', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const options: AddTableOptions = {
      name: 'Table valign lifecycle',
      valign: 'middle',
      columnWidths: inches(2),
      rowHeights: inches(1),
    };
    const table = slide.addTable([[
      'Inherited string',
      { text: 'Inherited object' },
      { text: 'Top override', options: { valign: 'top' } },
      { text: 'Bottom override', options: { valign: 'bottom' } },
    ]], options);
    expect(table.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual([
      'middle',
      'middle',
      'top',
      'bottom',
    ]);

    const sourceIndex = model.slides.indexOf(slide);
    const duplicate = model.duplicateSlide(sourceIndex);
    const duplicateTable = duplicate.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(duplicateTable).toBeInstanceOf(TableModel);
    expect(duplicateTable!.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual([
      'middle',
      'middle',
      'top',
      'bottom',
    ]);

    table.setCellVerticalAlignment(0, 0, undefined);
    table.setCellVerticalAlignment(0, 1, 'bottom');
    expect(table.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual([
      undefined,
      'bottom',
      'top',
      'bottom',
    ]);
    expect(duplicateTable!.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual([
      'middle',
      'middle',
      'top',
      'bottom',
    ]);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri);
    const reopenedTable = reopenedSlide?.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel
        && shape.name === 'Table valign lifecycle',
    );
    expect(reopenedTable).toBeInstanceOf(TableModel);
    expect(reopenedTable!.rows[0]!.cells.map(
      ({ verticalAlignment }) => verticalAlignment)).toEqual([
      undefined,
      'bottom',
      'top',
      'bottom',
    ]);
    expect(reopenedTable!.columnWidths).toEqual(Array(4).fill(inches(2)));
    expect(reopenedTable!.rowHeights).toEqual([inches(1)]);
  });

  it('projects and atomically edits uniform table borders', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const table = slide.addTable([
      [
        {
          text: 'North',
          options: {
            fill: {
              kind: 'solid',
              color: { kind: 'srgb', value: 'D9EAF7' },
              transparency: 25,
            },
            fit: 'shrink',
            margin: [1, 2, 3, 4],
            textDirection: 'vert270',
            valign: 'middle',
          },
        },
        'South',
      ],
      ['East', 'West'],
    ], {
      name: 'Uniform table borders',
      align: 'center',
      border: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent1' },
        width: 1.5,
        style: 'dash',
      },
      columnWidths: [inches(2), inches(3)],
      rowHeights: [inches(0.75), inches(1.25)],
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
    const nonBorderState = () => table.rows.map(({ cells }) => cells.map(({
      text,
      fill,
      horizontalAlignment,
      margins,
      textDirection,
      textFit,
      verticalAlignment,
    }) => ({
      text,
      fill,
      horizontalAlignment,
      margins,
      textDirection,
      textFit,
      verticalAlignment,
    })));
    const initialNonBorderState = nonBorderState();
    const initialTransform = table.transform;
    const initialColumnWidths = table.columnWidths;
    const initialRowHeights = table.rowHeights;
    const untouchedPartUri = '/ppt/custom/opaque1.bin';
    const untouchedBefore = pkg.requirePart(untouchedPartUri).bytes.slice();

    expect(table.borders).toEqual(initialBorders);
    const detached = table.borders!;
    const detachedTop = detached.top;
    expect(detachedTop?.kind).toBe('line');
    if (detachedTop?.kind === 'line') {
      (detachedTop.color as { value: string }).value = 'accent6';
      (detachedTop as { width: number }).width = 99;
    }
    expect(table.borders).toEqual(initialBorders);
    const noOpBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...pkg.mutations];
    table.borders = initialLine;
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(noOpBytes);
    expect(pkg.mutations).toEqual(noOpJournal);

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
    expect(nonBorderState()).toEqual(initialNonBorderState);

    table.borders = { kind: 'none' };
    const noneBorders = four({ kind: 'none' as const });
    expect(table.borders).toEqual(noneBorders);
    expect(table.rows.flatMap(({ cells }) => cells).map(({ borders }) => borders))
      .toEqual(Array(4).fill(noneBorders));

    table.borders = {};
    expect(table.borders).toBeUndefined();
    expect(table.rows.flatMap(({ cells }) => cells)
      .every(({ borders }) => borders === undefined)).toBe(true);
    const clearBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const clearJournal = [...pkg.mutations];
    table.borders = undefined;
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(clearBytes);
    expect(pkg.mutations).toEqual(clearJournal);

    table.borders = { kind: 'none' };
    const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
    const duplicateTable = duplicate.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(duplicateTable).toBeInstanceOf(TableModel);
    expect(duplicateTable!.borders).toEqual(noneBorders);
    const finalLine = {
      kind: 'line' as const,
      color: { kind: 'scheme' as const, value: 'accent2' as const },
      width: 0,
      style: 'solid' as const,
    };
    const finalBorders = four(finalLine);
    table.borders = finalLine;
    expect(table.borders).toEqual(finalBorders);
    expect(duplicateTable!.borders).toEqual(noneBorders);

    const beforeRollback = pkg.requirePart(slide.partUri).bytes.slice();
    const rollbackJournal = [...pkg.mutations];
    expect(() => pkg.transaction(() => {
      table.borders = { kind: 'none' };
      throw new Error('restore table-level borders');
    })).toThrow('restore table-level borders');
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(table.borders).toEqual(finalBorders);

    for (const invalid of [
      null,
      false,
      '',
      [],
      { type: 'solid', color: '4472C4', pt: 1 },
      { kind: 'line', color: { kind: 'srgb', value: 'bad' }, width: 1 },
      { kind: 'line', color: { kind: 'srgb', value: '4472C4' }, width: -1 },
      Object.create({ kind: 'none' }),
      Symbol('borders'),
    ]) {
      const beforeInvalid = pkg.requirePart(slide.partUri).bytes.slice();
      const invalidJournal = [...pkg.mutations];
      expect(() => {
        table.borders = invalid as never;
      }).toThrow();
      expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeInvalid);
      expect(pkg.mutations).toEqual(invalidJournal);
    }

    model.moveSlide(model.slides.indexOf(duplicate), 0);
    expect(table.borders).toEqual(finalBorders);
    expect(duplicateTable!.borders).toEqual(noneBorders);
    expect(nonBorderState()).toEqual(initialNonBorderState);
    expect(table.transform).toEqual(initialTransform);
    expect(table.columnWidths).toEqual(initialColumnWidths);
    expect(table.rowHeights).toEqual(initialRowHeights);
    expect(pkg.requirePart(untouchedPartUri).bytes).toEqual(untouchedBefore);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSourceTable = reopened.slides
      .find(({ partUri }) => partUri === slide.partUri)
      ?.shapes.find((shape): shape is TableModel => shape instanceof TableModel);
    const reopenedDuplicateTable = reopened.slides
      .find(({ partUri }) => partUri === duplicate.partUri)
      ?.shapes.find((shape): shape is TableModel => shape instanceof TableModel);
    expect(reopenedSourceTable).toBeInstanceOf(TableModel);
    expect(reopenedDuplicateTable).toBeInstanceOf(TableModel);
    expect(reopenedSourceTable!.borders).toEqual(finalBorders);
    expect(reopenedSourceTable!.rows.flatMap(({ cells }) => cells).map(({ borders }) => borders))
      .toEqual(Array(4).fill(finalBorders));
    expect(reopenedDuplicateTable!.borders).toEqual(noneBorders);
    expect(reopenedDuplicateTable!.rows.flatMap(({ cells }) => cells)
      .map(({ borders }) => borders)).toEqual(Array(4).fill(noneBorders));
    expect(reopenedSourceTable!.transform).toEqual(initialTransform);
    expect(reopenedSourceTable!.columnWidths).toEqual(initialColumnWidths);
    expect(reopenedSourceTable!.rowHeights).toEqual(initialRowHeights);
    expect(reopened.opcPackage.requirePart(untouchedPartUri).bytes).toEqual(untouchedBefore);

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

  it('rejects unsafe table-level border edits without partial package mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const table = slide.addTable([['First', 'Second']], {
      border: {
        kind: 'line',
        color: { kind: 'srgb', value: '4472C4' },
        width: 1,
      },
    });
    const original = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    const finalPropertiesEnd = original.lastIndexOf('</a:tcPr>');
    expect(finalPropertiesEnd).toBeGreaterThanOrEqual(0);
    const repeatedSide = original.slice(0, finalPropertiesEnd) +
      '<a:lnL w="0"><a:noFill/></a:lnL>' + original.slice(finalPropertiesEnd);
    pkg.setPart(slide.partUri, repeatedSide, pkg.requirePart(slide.partUri).contentType);
    expect(table.borders).toBeUndefined();
    const beforeUnsafe = pkg.requirePart(slide.partUri).bytes.slice();
    const unsafeJournal = [...pkg.mutations];
    expect(() => {
      table.borders = { kind: 'none' };
    }).toThrow(ModelParseError);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeUnsafe);
    expect(pkg.mutations).toEqual(unsafeJournal);

    const finalWidth = original.lastIndexOf('w="12700"');
    expect(finalWidth).toBeGreaterThanOrEqual(0);
    const malformed = original.slice(0, finalWidth) + 'w="invalid"' +
      original.slice(finalWidth + 'w="12700"'.length);
    pkg.setPart(slide.partUri, malformed, pkg.requirePart(slide.partUri).contentType);
    expect(table.borders).toBeUndefined();
    table.borders = { kind: 'none' };
    const noneBorders = {
      top: { kind: 'none' as const },
      right: { kind: 'none' as const },
      bottom: { kind: 'none' as const },
      left: { kind: 'none' as const },
    };
    expect(table.borders).toEqual(noneBorders);
    expect(table.rows[0]!.cells.map(({ borders }) => borders))
      .toEqual(Array(2).fill(noneBorders));
  });

  it('reopens table-level borders in every presentation format', async () => {
    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const model = new PresentationModel(await OpcPackage.open(
        await modelFixture(profile.presentationContentType),
      ));
      const table = model.addSlide().addTable([['A', 'B']], {
        border: { kind: 'none' },
      });
      table.borders = {
        top: {
          kind: 'line',
          color: { kind: 'srgb', value: 'D9EAF7' },
          width: 2,
          style: 'dash',
        },
        bottom: { kind: 'none' },
      };
      const reopened = new PresentationModel(await OpcPackage.open(await model.opcPackage.write()));
      expect(reopened.format).toBe(profile.format);
      const reopenedTable = reopened.slides.at(-1)!.shapes[0] as TableModel;
      const expected = {
        top: {
          kind: 'line' as const,
          color: { kind: 'srgb' as const, value: 'D9EAF7' },
          width: 2,
          style: 'dash' as const,
        },
        bottom: { kind: 'none' as const },
      };
      expect(reopenedTable.borders).toEqual(expected);
      expect(reopenedTable.rows[0]!.cells.map(({ borders }) => borders))
        .toEqual(Array(2).fill(expected));
    }
  });

  it('projects and atomically edits uniform table fill', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const table = slide.addTable([
      [
        {
          text: 'North',
          options: {
            border: {
              kind: 'line',
              color: { kind: 'scheme', value: 'accent2' },
              width: 1.5,
              style: 'dash',
            },
            fit: 'shrink',
            margin: [1, 2, 3, 4],
            textDirection: 'vert270',
            valign: 'middle',
          },
        },
        'South',
      ],
      ['East', 'West'],
    ], {
      name: 'Uniform table fill',
      align: 'center',
      fill: {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 25,
      },
      columnWidths: [inches(2), inches(3)],
      rowHeights: [inches(0.75), inches(1.25)],
    });
    const nonFillState = () => table.rows.map(({ cells }) => cells.map(({
      text,
      borders,
      horizontalAlignment,
      margins,
      textDirection,
      textFit,
      verticalAlignment,
    }) => ({
      text,
      borders,
      horizontalAlignment,
      margins,
      textDirection,
      textFit,
      verticalAlignment,
    })));
    const initialNonFillState = nonFillState();
    const initialTransform = table.transform;
    const initialColumnWidths = table.columnWidths;
    const initialRowHeights = table.rowHeights;
    const untouchedPartUri = '/ppt/custom/opaque1.bin';
    const untouchedBefore = pkg.requirePart(untouchedPartUri).bytes.slice();
    const initialFill = {
      kind: 'solid' as const,
      color: { kind: 'scheme' as const, value: 'accent1' as const },
      transparency: 25,
    };

    expect(table.fill).toEqual(initialFill);
    const detached = table.fill!;
    expect(detached.kind).toBe('solid');
    if (detached.kind === 'solid') {
      (detached.color as { value: string }).value = 'accent6';
    }
    (detached as { transparency?: number }).transparency = 99;
    expect(table.fill).toEqual(initialFill);
    const noOpBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...pkg.mutations];
    table.fill = {
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 25,
    };
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(noOpBytes);
    expect(pkg.mutations).toEqual(noOpJournal);

    table.setCellFill(0, 1, { kind: 'none' });
    expect(table.fill).toBeUndefined();
    table.fill = { kind: 'none' };
    expect(table.fill).toEqual({ kind: 'none' });
    expect(table.rows.flatMap(({ cells }) => cells).map(({ fill }) => fill))
      .toEqual(Array(4).fill({ kind: 'none' }));
    expect(nonFillState()).toEqual(initialNonFillState);

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
    expect(table.rows.flatMap(({ cells }) => cells)
      .every(({ fill }) => fill === undefined)).toBe(true);
    const clearBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const clearJournal = [...pkg.mutations];
    table.fill = undefined;
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(clearBytes);
    expect(pkg.mutations).toEqual(clearJournal);

    table.fill = { kind: 'none' };
    const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
    const duplicateTable = duplicate.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(duplicateTable).toBeInstanceOf(TableModel);
    expect(duplicateTable!.fill).toEqual({ kind: 'none' });
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
    expect(duplicateTable!.fill).toEqual({ kind: 'none' });

    const beforeRollback = pkg.requirePart(slide.partUri).bytes.slice();
    const rollbackJournal = [...pkg.mutations];
    expect(() => pkg.transaction(() => {
      table.fill = { kind: 'none' };
      throw new Error('restore table-level fill');
    })).toThrow('restore table-level fill');
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(table.fill).toEqual(finalFill);

    for (const [invalid, ErrorType] of [
      [null, TypeError],
      [false, TypeError],
      ['', TypeError],
      [{}, TypeError],
      [{ kind: 'solid' }, TypeError],
      [{ kind: 'solid', color: { kind: 'srgb', value: 'bad' } }, TypeError],
      [{
        kind: 'solid',
        color: { kind: 'srgb', value: 'D9EAF7' },
        transparency: 101,
      }, RangeError],
      [Object.create({ kind: 'none' }), TypeError],
      [Symbol('fill'), TypeError],
    ] as const) {
      const beforeInvalid = pkg.requirePart(slide.partUri).bytes.slice();
      const invalidJournal = [...pkg.mutations];
      expect(() => {
        table.fill = invalid as never;
      }).toThrow(ErrorType);
      expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeInvalid);
      expect(pkg.mutations).toEqual(invalidJournal);
    }

    model.moveSlide(model.slides.indexOf(duplicate), 0);
    expect(table.fill).toEqual(finalFill);
    expect(duplicateTable!.fill).toEqual({ kind: 'none' });
    expect(nonFillState()).toEqual(initialNonFillState);
    expect(table.transform).toEqual(initialTransform);
    expect(table.columnWidths).toEqual(initialColumnWidths);
    expect(table.rowHeights).toEqual(initialRowHeights);
    expect(pkg.requirePart(untouchedPartUri).bytes).toEqual(untouchedBefore);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSourceTable = reopened.slides
      .find(({ partUri }) => partUri === slide.partUri)
      ?.shapes.find((shape): shape is TableModel => shape instanceof TableModel);
    const reopenedDuplicateTable = reopened.slides
      .find(({ partUri }) => partUri === duplicate.partUri)
      ?.shapes.find((shape): shape is TableModel => shape instanceof TableModel);
    expect(reopenedSourceTable).toBeInstanceOf(TableModel);
    expect(reopenedDuplicateTable).toBeInstanceOf(TableModel);
    expect(reopenedSourceTable!.fill).toEqual(finalFill);
    expect(reopenedSourceTable!.rows.flatMap(({ cells }) => cells).map(({ fill }) => fill))
      .toEqual(Array(4).fill(finalFill));
    expect(reopenedDuplicateTable!.fill).toEqual({ kind: 'none' });
    expect(reopenedDuplicateTable!.rows.flatMap(({ cells }) => cells).map(({ fill }) => fill))
      .toEqual(Array(4).fill({ kind: 'none' }));
    expect(reopenedSourceTable!.transform).toEqual(initialTransform);
    expect(reopenedSourceTable!.columnWidths).toEqual(initialColumnWidths);
    expect(reopenedSourceTable!.rowHeights).toEqual(initialRowHeights);
    expect(reopened.opcPackage.requirePart(untouchedPartUri).bytes).toEqual(untouchedBefore);

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

  it('rejects unsafe table-level fill edits without partial package mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const table = slide.addTable([['First', 'Second']], {
      fill: { kind: 'solid', color: { kind: 'srgb', value: 'D9EAF7' } },
    });
    const original = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    const finalPropertiesEnd = original.lastIndexOf('</a:tcPr>');
    expect(finalPropertiesEnd).toBeGreaterThanOrEqual(0);
    const repeatedFill = original.slice(0, finalPropertiesEnd) +
      '<a:noFill/>' + original.slice(finalPropertiesEnd);
    pkg.setPart(slide.partUri, repeatedFill, pkg.requirePart(slide.partUri).contentType);
    expect(table.fill).toBeUndefined();
    const beforeUnsafe = pkg.requirePart(slide.partUri).bytes.slice();
    const unsafeJournal = [...pkg.mutations];
    expect(() => {
      table.fill = { kind: 'none' };
    }).toThrow(ModelParseError);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeUnsafe);
    expect(pkg.mutations).toEqual(unsafeJournal);

    const finalColor = original.lastIndexOf('val="D9EAF7"');
    expect(finalColor).toBeGreaterThanOrEqual(0);
    const malformed = original.slice(0, finalColor) + 'val="ZZZZZZ"' +
      original.slice(finalColor + 'val="D9EAF7"'.length);
    pkg.setPart(slide.partUri, malformed, pkg.requirePart(slide.partUri).contentType);
    expect(table.fill).toBeUndefined();
    table.fill = { kind: 'none' };
    expect(table.fill).toEqual({ kind: 'none' });
    expect(table.rows[0]!.cells.map(({ fill }) => fill))
      .toEqual(Array(2).fill({ kind: 'none' }));
  });

  it('reopens table-level fill in every presentation format', async () => {
    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const model = new PresentationModel(await OpcPackage.open(
        await modelFixture(profile.presentationContentType),
      ));
      const table = model.addSlide().addTable([['A', 'B']], {
        fill: { kind: 'none' },
      });
      table.fill = {
        kind: 'solid',
        color: { kind: 'srgb', value: 'D9EAF7' },
        transparency: 0,
      };
      const reopened = new PresentationModel(await OpcPackage.open(await model.opcPackage.write()));
      expect(reopened.format).toBe(profile.format);
      const reopenedTable = reopened.slides.at(-1)!.shapes[0] as TableModel;
      expect(reopenedTable.fill).toEqual({
        kind: 'solid',
        color: { kind: 'srgb', value: 'D9EAF7' },
        transparency: 0,
      });
      expect(reopenedTable.rows[0]!.cells.map(({ fill }) => fill)).toEqual(Array(2).fill({
        kind: 'solid',
        color: { kind: 'srgb', value: 'D9EAF7' },
        transparency: 0,
      }));
    }
  });

  it('projects and atomically edits uniform table margins', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const table = slide.addTable([
      [
        {
          text: 'North',
          options: {
            border: {
              kind: 'line',
              color: { kind: 'scheme', value: 'accent1' },
              width: 1.5,
              style: 'dash',
            },
            fill: {
              kind: 'solid',
              color: { kind: 'srgb', value: 'D9EAF7' },
              transparency: 25,
            },
            fit: 'shrink',
            textDirection: 'vert270',
            valign: 'middle',
          },
        },
        'South',
      ],
      ['East', 'West'],
    ], {
      name: 'Uniform table margins',
      align: 'center',
      margin: [3.6, 7.2, 10.8, 14.4],
      columnWidths: [inches(2), inches(3)],
      rowHeights: [inches(0.75), inches(1.25)],
    });
    const nonMarginState = () => table.rows.map(({ cells }) => cells.map(({
      text,
      borders,
      fill,
      horizontalAlignment,
      textDirection,
      textFit,
      verticalAlignment,
    }) => ({
      text,
      borders,
      fill,
      horizontalAlignment,
      textDirection,
      textFit,
      verticalAlignment,
    })));
    const initialNonMarginState = nonMarginState();
    const initialTransform = table.transform;
    const initialColumnWidths = table.columnWidths;
    const initialRowHeights = table.rowHeights;
    const untouchedPartUri = '/ppt/custom/opaque1.bin';
    const untouchedBefore = pkg.requirePart(untouchedPartUri).bytes.slice();

    expect(table.margins).toEqual({ top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 });
    const detached = table.margins!;
    (detached as { top?: number }).top = 99;
    expect(table.margins).toEqual({ top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 });
    const noOpBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...pkg.mutations];
    table.margins = [3.6, 7.2, 10.8, 14.4];
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(noOpBytes);
    expect(pkg.mutations).toEqual(noOpJournal);

    table.setCellMargins(0, 1, { top: 9 });
    expect(table.margins).toBeUndefined();
    table.margins = 6;
    expect(table.margins).toEqual({ top: 6, right: 6, bottom: 6, left: 6 });
    expect(table.rows.flatMap(({ cells }) => cells).map(({ margins }) => margins))
      .toEqual(Array(4).fill({ top: 6, right: 6, bottom: 6, left: 6 }));
    expect(nonMarginState()).toEqual(initialNonMarginState);

    table.margins = { top: 2, left: 4 };
    expect(table.margins).toEqual({ top: 2, left: 4 });
    expect(table.rows.flatMap(({ cells }) => cells).map(({ margins }) => margins))
      .toEqual(Array(4).fill({ top: 2, left: 4 }));
    table.margins = {};
    expect(table.margins).toBeUndefined();
    expect(table.rows.flatMap(({ cells }) => cells)
      .every(({ margins }) => margins === undefined)).toBe(true);
    const clearBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const clearJournal = [...pkg.mutations];
    table.margins = undefined;
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(clearBytes);
    expect(pkg.mutations).toEqual(clearJournal);

    table.margins = [1, 2, 3, 4];
    const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
    const duplicateTable = duplicate.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(duplicateTable).toBeInstanceOf(TableModel);
    expect(duplicateTable!.margins).toEqual({ top: 1, right: 2, bottom: 3, left: 4 });
    table.margins = 5;
    expect(table.margins).toEqual({ top: 5, right: 5, bottom: 5, left: 5 });
    expect(duplicateTable!.margins).toEqual({ top: 1, right: 2, bottom: 3, left: 4 });
    table.margins = [1, 2, 3, 4];

    const beforeRollback = pkg.requirePart(slide.partUri).bytes.slice();
    const rollbackJournal = [...pkg.mutations];
    expect(() => pkg.transaction(() => {
      table.margins = { bottom: 9 };
      throw new Error('restore table-level margins');
    })).toThrow('restore table-level margins');
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(table.margins).toEqual({ top: 1, right: 2, bottom: 3, left: 4 });

    for (const [invalid, ErrorType] of [
      [null, TypeError],
      [false, TypeError],
      [true, TypeError],
      ['', TypeError],
      [[1, 2, 3], RangeError],
      [[1, 2, 3, 4, 5], RangeError],
      [{ middle: 1 }, TypeError],
      [{ top: Number.NaN }, TypeError],
      [Object.create({ top: 1 }), TypeError],
      [Symbol('margins'), TypeError],
    ] as const) {
      const beforeInvalid = pkg.requirePart(slide.partUri).bytes.slice();
      const invalidJournal = [...pkg.mutations];
      expect(() => {
        table.margins = invalid as never;
      }).toThrow(ErrorType);
      expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeInvalid);
      expect(pkg.mutations).toEqual(invalidJournal);
    }

    model.moveSlide(model.slides.indexOf(duplicate), 0);
    expect(table.margins).toEqual({ top: 1, right: 2, bottom: 3, left: 4 });
    expect(duplicateTable!.margins).toEqual({ top: 1, right: 2, bottom: 3, left: 4 });
    expect(nonMarginState()).toEqual(initialNonMarginState);
    expect(table.transform).toEqual(initialTransform);
    expect(table.columnWidths).toEqual(initialColumnWidths);
    expect(table.rowHeights).toEqual(initialRowHeights);
    expect(pkg.requirePart(untouchedPartUri).bytes).toEqual(untouchedBefore);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSourceTable = reopened.slides
      .find(({ partUri }) => partUri === slide.partUri)
      ?.shapes.find((shape): shape is TableModel => shape instanceof TableModel);
    const reopenedDuplicateTable = reopened.slides
      .find(({ partUri }) => partUri === duplicate.partUri)
      ?.shapes.find((shape): shape is TableModel => shape instanceof TableModel);
    expect(reopenedSourceTable).toBeInstanceOf(TableModel);
    expect(reopenedDuplicateTable).toBeInstanceOf(TableModel);
    for (const reopenedTable of [reopenedSourceTable!, reopenedDuplicateTable!]) {
      expect(reopenedTable.margins).toEqual({ top: 1, right: 2, bottom: 3, left: 4 });
      expect(reopenedTable.rows.flatMap(({ cells }) => cells).map(({ margins }) => margins))
        .toEqual(Array(4).fill({ top: 1, right: 2, bottom: 3, left: 4 }));
    }
    expect(reopenedSourceTable!.transform).toEqual(initialTransform);
    expect(reopenedSourceTable!.columnWidths).toEqual(initialColumnWidths);
    expect(reopenedSourceTable!.rowHeights).toEqual(initialRowHeights);
    expect(reopened.opcPackage.requirePart(untouchedPartUri).bytes).toEqual(untouchedBefore);

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

  it('rejects unsafe table-level margin edits without partial package mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const table = slide.addTable([['First', 'Second']], { margin: 1 });
    const original = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    const finalPropertiesOffset = original.lastIndexOf('<a:tcPr');
    const finalPropertiesEnd = original.indexOf('>', finalPropertiesOffset);
    expect(finalPropertiesOffset).toBeGreaterThanOrEqual(0);
    expect(finalPropertiesEnd).toBeGreaterThan(finalPropertiesOffset);
    const repeatedMargin = original.slice(0, finalPropertiesEnd) +
      ' marT="25400"' + original.slice(finalPropertiesEnd);
    pkg.setPart(slide.partUri, repeatedMargin, pkg.requirePart(slide.partUri).contentType);
    expect(table.margins).toBeUndefined();
    const beforeUnsafe = pkg.requirePart(slide.partUri).bytes.slice();
    const unsafeJournal = [...pkg.mutations];
    expect(() => {
      table.margins = 2;
    }).toThrow(ModelParseError);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeUnsafe);
    expect(pkg.mutations).toEqual(unsafeJournal);

    const malformed = original.replace(/marT="12700"(?=[^<]*>[^]*$)/, 'marT="bad"');
    pkg.setPart(slide.partUri, malformed, pkg.requirePart(slide.partUri).contentType);
    expect(table.margins).toBeUndefined();
    table.margins = 2;
    expect(table.margins).toEqual({ top: 2, right: 2, bottom: 2, left: 2 });
  });

  it('reopens table-level margins in every presentation format', async () => {
    for (const profile of Object.values(PRESENTATION_FORMAT_PROFILES)) {
      const model = new PresentationModel(await OpcPackage.open(
        await modelFixture(profile.presentationContentType),
      ));
      const table = model.addSlide().addTable([['A', 'B']], { margin: 1 });
      table.margins = { top: 2, left: 4 };
      const reopened = new PresentationModel(await OpcPackage.open(await model.opcPackage.write()));
      expect(reopened.format).toBe(profile.format);
      const reopenedTable = reopened.slides.at(-1)!.shapes[0] as TableModel;
      expect(reopenedTable.margins).toEqual({ top: 2, left: 4 });
      expect(reopenedTable.rows[0]!.cells.map(({ margins }) => margins))
        .toEqual(Array(2).fill({ top: 2, left: 4 }));
    }
  });

  it('projects and atomically edits uniform table horizontal alignment', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const table = slide.addTable([
      [
        {
          text: 'North',
          options: {
            border: {
              kind: 'line',
              color: { kind: 'scheme', value: 'accent1' },
              width: 1.5,
              style: 'dash',
            },
            fill: {
              kind: 'solid',
              color: { kind: 'srgb', value: 'D9EAF7' },
              transparency: 25,
            },
            fit: 'shrink',
            margin: { top: 5, left: 8 },
            textDirection: 'vert270',
            valign: 'middle',
          },
        },
        'South',
      ],
      ['East', 'West'],
    ], {
      name: 'Uniform table horizontal alignment',
      align: 'center',
      columnWidths: [inches(2), inches(3)],
      rowHeights: [inches(0.75), inches(1.25)],
    });
    const nonAlignmentState = () => table.rows.map(({ cells }) => cells.map(({
      text,
      borders,
      fill,
      margins,
      textDirection,
      textFit,
      verticalAlignment,
    }) => ({
      text,
      borders,
      fill,
      margins,
      textDirection,
      textFit,
      verticalAlignment,
    })));
    const initialNonAlignmentState = nonAlignmentState();
    const initialTransform = table.transform;
    const initialColumnWidths = table.columnWidths;
    const initialRowHeights = table.rowHeights;
    const untouchedPartUri = '/ppt/custom/opaque1.bin';
    const untouchedBefore = pkg.requirePart(untouchedPartUri).bytes.slice();

    expect(table.horizontalAlignment).toBe('center');
    const noOpBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...pkg.mutations];
    void table.horizontalAlignment;
    table.horizontalAlignment = 'center';
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(noOpBytes);
    expect(pkg.mutations).toEqual(noOpJournal);

    table.setCellHorizontalAlignment(0, 1, 'right');
    expect(table.horizontalAlignment).toBeUndefined();
    table.horizontalAlignment = 'justify';
    expect(table.horizontalAlignment).toBe('justify');
    expect(table.rows.flatMap(({ cells }) => cells)
      .map(({ horizontalAlignment }) => horizontalAlignment))
      .toEqual(['justify', 'justify', 'justify', 'justify']);
    expect(nonAlignmentState()).toEqual(initialNonAlignmentState);

    table.horizontalAlignment = 'left';
    expect(table.horizontalAlignment).toBe('left');
    const leftXml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(leftXml.match(/<a:pPr\b[^>]* algn="l"/g)).toHaveLength(4);
    const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
    const duplicateTable = duplicate.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(duplicateTable).toBeInstanceOf(TableModel);
    expect(duplicateTable!.horizontalAlignment).toBe('left');

    table.horizontalAlignment = undefined;
    expect(table.horizontalAlignment).toBeUndefined();
    expect(table.rows.flatMap(({ cells }) => cells)
      .every(({ horizontalAlignment }) => horizontalAlignment === undefined)).toBe(true);
    expect(duplicateTable!.horizontalAlignment).toBe('left');
    table.horizontalAlignment = 'right';
    expect(table.horizontalAlignment).toBe('right');

    const beforeRollback = pkg.requirePart(slide.partUri).bytes.slice();
    const rollbackJournal = [...pkg.mutations];
    expect(() => pkg.transaction(() => {
      table.horizontalAlignment = 'center';
      throw new Error('restore table-level horizontal alignment');
    })).toThrow('restore table-level horizontal alignment');
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(table.horizontalAlignment).toBe('right');

    for (const invalid of [
      null,
      false,
      true,
      0,
      '',
      'l',
      'ctr',
      'dist',
      'Left',
      ' left',
      [],
      {},
      Symbol('align'),
    ]) {
      const beforeInvalid = pkg.requirePart(slide.partUri).bytes.slice();
      const invalidJournal = [...pkg.mutations];
      expect(() => {
        table.horizontalAlignment = invalid as never;
      }, String(invalid)).toThrow(TypeError);
      expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeInvalid);
      expect(pkg.mutations).toEqual(invalidJournal);
    }

    model.moveSlide(model.slides.indexOf(duplicate), 0);
    expect(table.horizontalAlignment).toBe('right');
    expect(duplicateTable!.horizontalAlignment).toBe('left');
    expect(nonAlignmentState()).toEqual(initialNonAlignmentState);
    expect(table.transform).toEqual(initialTransform);
    expect(table.columnWidths).toEqual(initialColumnWidths);
    expect(table.rowHeights).toEqual(initialRowHeights);
    expect(pkg.requirePart(untouchedPartUri).bytes).toEqual(untouchedBefore);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSourceTable = reopened.slides
      .find(({ partUri }) => partUri === slide.partUri)
      ?.shapes.find((shape): shape is TableModel => shape instanceof TableModel);
    const reopenedDuplicateTable = reopened.slides
      .find(({ partUri }) => partUri === duplicate.partUri)
      ?.shapes.find((shape): shape is TableModel => shape instanceof TableModel);
    expect(reopenedSourceTable).toBeInstanceOf(TableModel);
    expect(reopenedDuplicateTable).toBeInstanceOf(TableModel);
    expect(reopenedSourceTable!.horizontalAlignment).toBe('right');
    expect(reopenedDuplicateTable!.horizontalAlignment).toBe('left');
    expect(reopenedSourceTable!.rows.flatMap(({ cells }) => cells)
      .map(({ horizontalAlignment }) => horizontalAlignment))
      .toEqual(['right', 'right', 'right', 'right']);
    expect(reopenedDuplicateTable!.rows.flatMap(({ cells }) => cells)
      .map(({ horizontalAlignment }) => horizontalAlignment))
      .toEqual(['left', 'left', 'left', 'left']);
    expect(reopenedSourceTable!.transform).toEqual(initialTransform);
    expect(reopenedSourceTable!.columnWidths).toEqual(initialColumnWidths);
    expect(reopenedSourceTable!.rowHeights).toEqual(initialRowHeights);
    expect(reopened.opcPackage.requirePart(untouchedPartUri).bytes).toEqual(untouchedBefore);

    if (false) {
      const alignment: TextAlignment | undefined = table.horizontalAlignment;
      table.horizontalAlignment = 'left';
      table.horizontalAlignment = 'center';
      table.horizontalAlignment = 'right';
      table.horizontalAlignment = 'justify';
      table.horizontalAlignment = undefined;
      // @ts-expect-error unsupported table horizontal alignment
      table.horizontalAlignment = 'dist';
      void alignment;
    }
  });

  it('rejects unsafe table-level horizontal alignment edits without partial mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const table = slide.addTable([['First', 'Second']], { align: 'center' });
    const original = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    const secondTextOffset = original.indexOf('>Second</a:t>');
    const secondPropertiesOffset = original.lastIndexOf('<a:pPr', secondTextOffset);
    const secondPropertiesEnd = original.indexOf('>', secondPropertiesOffset);
    expect(secondTextOffset).toBeGreaterThanOrEqual(0);
    expect(secondPropertiesOffset).toBeGreaterThanOrEqual(0);
    expect(secondPropertiesEnd).toBeGreaterThan(secondPropertiesOffset);
    const repeatedAlignment = original.slice(0, secondPropertiesEnd) +
      ' algn="r"' + original.slice(secondPropertiesEnd);
    pkg.setPart(
      slide.partUri,
      repeatedAlignment,
      pkg.requirePart(slide.partUri).contentType,
    );
    expect(table.horizontalAlignment).toBeUndefined();
    const beforeUnsafe = pkg.requirePart(slide.partUri).bytes.slice();
    const unsafeJournal = [...pkg.mutations];
    expect(() => {
      table.horizontalAlignment = 'justify';
    }).toThrow(ModelParseError);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeUnsafe);
    expect(pkg.mutations).toEqual(unsafeJournal);

    const secondParagraphStart = original.lastIndexOf('<a:p>', secondTextOffset);
    const secondParagraphEnd = original.indexOf('</a:p>', secondTextOffset) + '</a:p>'.length;
    expect(secondParagraphStart).toBeGreaterThanOrEqual(0);
    expect(secondParagraphEnd).toBeGreaterThan(secondTextOffset);
    const secondParagraph = original.slice(secondParagraphStart, secondParagraphEnd);
    const multipleParagraphs = original.slice(0, secondParagraphEnd) + secondParagraph +
      original.slice(secondParagraphEnd);
    pkg.setPart(
      slide.partUri,
      multipleParagraphs,
      pkg.requirePart(slide.partUri).contentType,
    );
    expect(table.horizontalAlignment).toBeUndefined();
    const beforeMultiple = pkg.requirePart(slide.partUri).bytes.slice();
    const multipleJournal = [...pkg.mutations];
    expect(() => {
      table.horizontalAlignment = 'left';
    }).toThrow(ModelParseError);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeMultiple);
    expect(pkg.mutations).toEqual(multipleJournal);
  });

  it('projects and atomically edits uniform table vertical alignment', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const table = slide.addTable([
      [
        {
          text: 'North',
          options: {
            align: 'center',
            border: {
              kind: 'line',
              color: { kind: 'scheme', value: 'accent1' },
              width: 1.5,
              style: 'dash',
            },
            fill: {
              kind: 'solid',
              color: { kind: 'srgb', value: 'D9EAF7' },
              transparency: 25,
            },
            fit: 'shrink',
            margin: { top: 5, left: 8 },
            textDirection: 'vert270',
          },
        },
        'South',
      ],
      ['East', 'West'],
    ], {
      name: 'Uniform table vertical alignment',
      valign: 'middle',
      columnWidths: [inches(2), inches(3)],
      rowHeights: [inches(0.75), inches(1.25)],
    });
    const nonAlignmentState = () => table.rows.map(({ cells }) => cells.map(({
      text,
      borders,
      fill,
      horizontalAlignment,
      margins,
      textDirection,
      textFit,
    }) => ({
      text,
      borders,
      fill,
      horizontalAlignment,
      margins,
      textDirection,
      textFit,
    })));
    const initialNonAlignmentState = nonAlignmentState();
    const initialTransform = table.transform;
    const initialColumnWidths = table.columnWidths;
    const initialRowHeights = table.rowHeights;
    const untouchedPartUri = '/ppt/custom/opaque1.bin';
    const untouchedBefore = pkg.requirePart(untouchedPartUri).bytes.slice();

    expect(table.verticalAlignment).toBe('middle');
    const noOpBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...pkg.mutations];
    void table.verticalAlignment;
    table.verticalAlignment = 'middle';
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(noOpBytes);
    expect(pkg.mutations).toEqual(noOpJournal);

    table.setCellVerticalAlignment(0, 1, 'top');
    expect(table.verticalAlignment).toBeUndefined();
    table.verticalAlignment = 'bottom';
    expect(table.verticalAlignment).toBe('bottom');
    expect(table.rows.flatMap(({ cells }) => cells)
      .map(({ verticalAlignment }) => verticalAlignment))
      .toEqual(['bottom', 'bottom', 'bottom', 'bottom']);
    expect(nonAlignmentState()).toEqual(initialNonAlignmentState);

    const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
    const duplicateTable = duplicate.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(duplicateTable).toBeInstanceOf(TableModel);
    expect(duplicateTable!.verticalAlignment).toBe('bottom');

    table.verticalAlignment = undefined;
    expect(table.verticalAlignment).toBeUndefined();
    expect(table.rows.flatMap(({ cells }) => cells)
      .every(({ verticalAlignment }) => verticalAlignment === undefined)).toBe(true);
    expect(duplicateTable!.verticalAlignment).toBe('bottom');
    table.verticalAlignment = 'top';
    expect(table.verticalAlignment).toBe('top');
    expect(duplicateTable!.verticalAlignment).toBe('bottom');

    const beforeRollback = pkg.requirePart(slide.partUri).bytes.slice();
    const rollbackJournal = [...pkg.mutations];
    expect(() => pkg.transaction(() => {
      table.verticalAlignment = 'middle';
      throw new Error('restore table-level vertical alignment');
    })).toThrow('restore table-level vertical alignment');
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(table.verticalAlignment).toBe('top');

    for (const invalid of [
      null,
      false,
      true,
      0,
      '',
      'Top',
      ' top',
      't',
      'ctr',
      'distributed',
      [],
      {},
      Symbol('top'),
    ]) {
      const beforeInvalid = pkg.requirePart(slide.partUri).bytes.slice();
      const invalidJournal = [...pkg.mutations];
      expect(() => {
        table.verticalAlignment = invalid as never;
      }, String(invalid)).toThrow(TypeError);
      expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeInvalid);
      expect(pkg.mutations).toEqual(invalidJournal);
    }

    model.moveSlide(model.slides.indexOf(duplicate), 0);
    expect(table.verticalAlignment).toBe('top');
    expect(duplicateTable!.verticalAlignment).toBe('bottom');
    expect(nonAlignmentState()).toEqual(initialNonAlignmentState);
    expect(table.transform).toEqual(initialTransform);
    expect(table.columnWidths).toEqual(initialColumnWidths);
    expect(table.rowHeights).toEqual(initialRowHeights);
    expect(pkg.requirePart(untouchedPartUri).bytes).toEqual(untouchedBefore);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSourceTable = reopened.slides
      .find(({ partUri }) => partUri === slide.partUri)
      ?.shapes.find((shape): shape is TableModel => shape instanceof TableModel);
    const reopenedDuplicateTable = reopened.slides
      .find(({ partUri }) => partUri === duplicate.partUri)
      ?.shapes.find((shape): shape is TableModel => shape instanceof TableModel);
    expect(reopenedSourceTable).toBeInstanceOf(TableModel);
    expect(reopenedDuplicateTable).toBeInstanceOf(TableModel);
    expect(reopenedSourceTable!.verticalAlignment).toBe('top');
    expect(reopenedDuplicateTable!.verticalAlignment).toBe('bottom');
    expect(reopenedSourceTable!.rows.flatMap(({ cells }) => cells)
      .map(({ verticalAlignment }) => verticalAlignment))
      .toEqual(['top', 'top', 'top', 'top']);
    expect(reopenedDuplicateTable!.rows.flatMap(({ cells }) => cells)
      .map(({ verticalAlignment }) => verticalAlignment))
      .toEqual(['bottom', 'bottom', 'bottom', 'bottom']);
    expect(reopenedSourceTable!.transform).toEqual(initialTransform);
    expect(reopenedSourceTable!.columnWidths).toEqual(initialColumnWidths);
    expect(reopenedSourceTable!.rowHeights).toEqual(initialRowHeights);
    expect(reopened.opcPackage.requirePart(untouchedPartUri).bytes).toEqual(untouchedBefore);

    if (false) {
      const alignment: TextBoxVerticalAlignment | undefined = table.verticalAlignment;
      table.verticalAlignment = 'top';
      table.verticalAlignment = 'middle';
      table.verticalAlignment = 'bottom';
      table.verticalAlignment = undefined;
      // @ts-expect-error unsupported table vertical alignment
      table.verticalAlignment = 'distributed';
      void alignment;
    }
  });

  it('rejects unsafe table-level vertical alignment edits without partial package mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const table = slide.addTable([['First', 'Second']], { valign: 'middle' });
    const original = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    const secondTextOffset = original.indexOf('>Second</a:t>');
    const secondPropertiesOffset = original.indexOf('<a:tcPr', secondTextOffset);
    const secondPropertiesEnd = original.indexOf('>', secondPropertiesOffset);
    expect(secondTextOffset).toBeGreaterThanOrEqual(0);
    expect(secondPropertiesOffset).toBeGreaterThan(secondTextOffset);
    expect(secondPropertiesEnd).toBeGreaterThan(secondPropertiesOffset);
    const repeatedAnchor = original.slice(0, secondPropertiesEnd) +
      ' anchor="b"' + original.slice(secondPropertiesEnd);
    pkg.setPart(
      slide.partUri,
      repeatedAnchor,
      pkg.requirePart(slide.partUri).contentType,
    );
    expect(table.verticalAlignment).toBeUndefined();
    const beforeUnsafe = pkg.requirePart(slide.partUri).bytes.slice();
    const unsafeJournal = [...pkg.mutations];
    expect(() => {
      table.verticalAlignment = 'bottom';
    }).toThrow(ModelParseError);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeUnsafe);
    expect(pkg.mutations).toEqual(unsafeJournal);

    const emptyTable = original.replace(/<a:tr\b[\s\S]*?<\/a:tr>/, '');
    pkg.setPart(
      slide.partUri,
      emptyTable,
      pkg.requirePart(slide.partUri).contentType,
    );
    expect(table.verticalAlignment).toBeUndefined();
    const beforeEmpty = pkg.requirePart(slide.partUri).bytes.slice();
    const emptyJournal = [...pkg.mutations];
    expect(() => {
      table.verticalAlignment = 'top';
    }).toThrow(ModelParseError);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeEmpty);
    expect(pkg.mutations).toEqual(emptyJournal);
  });

  it('projects and atomically edits uniform table text direction', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const table = slide.addTable([
      [
        {
          text: 'North',
          options: {
            align: 'center',
            border: {
              kind: 'line',
              color: { kind: 'scheme', value: 'accent1' },
              width: 1.5,
              style: 'dash',
            },
            fill: {
              kind: 'solid',
              color: { kind: 'srgb', value: 'D9EAF7' },
              transparency: 25,
            },
            fit: 'shrink',
            margin: { top: 5, left: 8 },
            valign: 'middle',
          },
        },
        'South',
      ],
      ['East', 'West'],
    ], {
      name: 'Uniform table text direction',
      textDirection: 'vert270',
      columnWidths: [inches(2), inches(3)],
      rowHeights: [inches(0.75), inches(1.25)],
    });
    const nonDirectionState = () => table.rows.map(({ cells }) => cells.map(({
      text,
      borders,
      fill,
      horizontalAlignment,
      margins,
      textFit,
      verticalAlignment,
    }) => ({
      text,
      borders,
      fill,
      horizontalAlignment,
      margins,
      textFit,
      verticalAlignment,
    })));
    const initialNonDirectionState = nonDirectionState();
    const initialTransform = table.transform;
    const initialColumnWidths = table.columnWidths;
    const initialRowHeights = table.rowHeights;
    const untouchedPartUri = '/ppt/custom/opaque1.bin';
    const untouchedBefore = pkg.requirePart(untouchedPartUri).bytes.slice();

    expect(table.textDirection).toBe('vert270');
    const noOpBytes = pkg.requirePart(slide.partUri).bytes.slice();
    const noOpJournal = [...pkg.mutations];
    void table.textDirection;
    table.textDirection = 'vert270';
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(noOpBytes);
    expect(pkg.mutations).toEqual(noOpJournal);

    table.setCellTextDirection(0, 1, 'vert');
    expect(table.textDirection).toBeUndefined();
    table.textDirection = 'wordArtVert';
    expect(table.textDirection).toBe('wordArtVert');
    expect(table.rows.flatMap(({ cells }) => cells)
      .map(({ textDirection }) => textDirection))
      .toEqual(['wordArtVert', 'wordArtVert', 'wordArtVert', 'wordArtVert']);
    expect(nonDirectionState()).toEqual(initialNonDirectionState);

    table.textDirection = 'horz';
    expect(table.textDirection).toBe('horz');
    expect(table.rows.flatMap(({ cells }) => cells)
      .every(({ textDirection }) => textDirection === 'horz')).toBe(true);
    const horizontalXml = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(horizontalXml.match(/<a:tcPr[^>]* vert="horz"/g)).toHaveLength(4);

    table.textDirection = 'wordArtVert';
    const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
    const duplicateTable = duplicate.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(duplicateTable).toBeInstanceOf(TableModel);
    expect(duplicateTable!.textDirection).toBe('wordArtVert');

    table.textDirection = undefined;
    expect(table.textDirection).toBeUndefined();
    expect(table.rows.flatMap(({ cells }) => cells)
      .every(({ textDirection }) => textDirection === undefined)).toBe(true);
    expect(duplicateTable!.textDirection).toBe('wordArtVert');
    table.textDirection = 'vert';
    expect(table.textDirection).toBe('vert');
    expect(duplicateTable!.textDirection).toBe('wordArtVert');

    const beforeRollback = pkg.requirePart(slide.partUri).bytes.slice();
    const rollbackJournal = [...pkg.mutations];
    expect(() => pkg.transaction(() => {
      table.textDirection = 'vert270';
      throw new Error('restore table-level text direction');
    })).toThrow('restore table-level text direction');
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(table.textDirection).toBe('vert');

    for (const invalid of [
      null,
      false,
      true,
      0,
      '',
      'horizontal',
      'Horz',
      ' vert',
      'vert90',
      'eaVert',
      [],
      {},
      Symbol('vert'),
    ]) {
      const beforeInvalid = pkg.requirePart(slide.partUri).bytes.slice();
      const invalidJournal = [...pkg.mutations];
      expect(() => {
        table.textDirection = invalid as never;
      }, String(invalid)).toThrow(TypeError);
      expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeInvalid);
      expect(pkg.mutations).toEqual(invalidJournal);
    }

    model.moveSlide(model.slides.indexOf(duplicate), 0);
    expect(table.textDirection).toBe('vert');
    expect(duplicateTable!.textDirection).toBe('wordArtVert');
    expect(nonDirectionState()).toEqual(initialNonDirectionState);
    expect(table.transform).toEqual(initialTransform);
    expect(table.columnWidths).toEqual(initialColumnWidths);
    expect(table.rowHeights).toEqual(initialRowHeights);
    expect(pkg.requirePart(untouchedPartUri).bytes).toEqual(untouchedBefore);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSourceTable = reopened.slides
      .find(({ partUri }) => partUri === slide.partUri)
      ?.shapes.find((shape): shape is TableModel => shape instanceof TableModel);
    const reopenedDuplicateTable = reopened.slides
      .find(({ partUri }) => partUri === duplicate.partUri)
      ?.shapes.find((shape): shape is TableModel => shape instanceof TableModel);
    expect(reopenedSourceTable).toBeInstanceOf(TableModel);
    expect(reopenedDuplicateTable).toBeInstanceOf(TableModel);
    expect(reopenedSourceTable!.textDirection).toBe('vert');
    expect(reopenedDuplicateTable!.textDirection).toBe('wordArtVert');
    expect(reopenedSourceTable!.rows.flatMap(({ cells }) => cells)
      .map(({ textDirection }) => textDirection))
      .toEqual(['vert', 'vert', 'vert', 'vert']);
    expect(reopenedDuplicateTable!.rows.flatMap(({ cells }) => cells)
      .map(({ textDirection }) => textDirection))
      .toEqual(['wordArtVert', 'wordArtVert', 'wordArtVert', 'wordArtVert']);
    expect(reopenedSourceTable!.transform).toEqual(initialTransform);
    expect(reopenedSourceTable!.columnWidths).toEqual(initialColumnWidths);
    expect(reopenedSourceTable!.rowHeights).toEqual(initialRowHeights);
    expect(reopened.opcPackage.requirePart(untouchedPartUri).bytes).toEqual(untouchedBefore);

    if (false) {
      const direction: TableCellTextDirection | undefined = table.textDirection;
      table.textDirection = 'horz';
      table.textDirection = 'vert';
      table.textDirection = 'vert270';
      table.textDirection = 'wordArtVert';
      table.textDirection = undefined;
      // @ts-expect-error unsupported table text direction
      table.textDirection = 'eaVert';
      void direction;
    }
  });

  it('rejects unsafe table-level text direction edits without partial package mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const table = slide.addTable([['First', 'Second']], { textDirection: 'vert270' });
    const original = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    const secondTextOffset = original.indexOf('>Second</a:t>');
    const secondPropertiesOffset = original.indexOf('<a:tcPr', secondTextOffset);
    const secondPropertiesEnd = original.indexOf('>', secondPropertiesOffset);
    expect(secondTextOffset).toBeGreaterThanOrEqual(0);
    expect(secondPropertiesOffset).toBeGreaterThan(secondTextOffset);
    expect(secondPropertiesEnd).toBeGreaterThan(secondPropertiesOffset);
    const repeatedDirection = original.slice(0, secondPropertiesEnd) +
      ' vert="vert"' + original.slice(secondPropertiesEnd);
    pkg.setPart(
      slide.partUri,
      repeatedDirection,
      pkg.requirePart(slide.partUri).contentType,
    );
    expect(table.textDirection).toBeUndefined();
    const beforeUnsafe = pkg.requirePart(slide.partUri).bytes.slice();
    const unsafeJournal = [...pkg.mutations];
    expect(() => {
      table.textDirection = 'wordArtVert';
    }).toThrow(ModelParseError);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeUnsafe);
    expect(pkg.mutations).toEqual(unsafeJournal);

    const emptyTable = original.replace(/<a:tr\b[\s\S]*?<\/a:tr>/, '');
    pkg.setPart(
      slide.partUri,
      emptyTable,
      pkg.requirePart(slide.partUri).contentType,
    );
    expect(table.textDirection).toBeUndefined();
    const beforeEmpty = pkg.requirePart(slide.partUri).bytes.slice();
    const emptyJournal = [...pkg.mutations];
    expect(() => {
      table.textDirection = 'vert';
    }).toThrow(ModelParseError);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeEmpty);
    expect(pkg.mutations).toEqual(emptyJournal);
  });

  it('materializes table text direction through edit, duplicate, write, and reopen', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const options: AddTableOptions = {
      name: 'Table text direction lifecycle',
      textDirection: 'vert270',
      columnWidths: inches(2),
      rowHeights: inches(1),
    };
    const table = slide.addTable([[
      'Inherited string',
      { text: 'Inherited object' },
      { text: 'Horizontal override', options: { textDirection: 'horz' } },
      { text: 'Vertical override', options: { textDirection: 'vert' } },
      { text: 'Stacked override', options: { textDirection: 'wordArtVert' } },
    ]], options);
    const materializedDirections = [
      'vert270',
      'vert270',
      undefined,
      'vert',
      'wordArtVert',
    ];
    expect(table.rows[0]!.cells.map(({ textDirection }) => textDirection))
      .toEqual(materializedDirections);

    const duplicate = model.duplicateSlide(model.slides.indexOf(slide));
    const duplicateTable = duplicate.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(duplicateTable).toBeInstanceOf(TableModel);
    expect(duplicateTable!.rows[0]!.cells.map(
      ({ textDirection }) => textDirection)).toEqual(materializedDirections);

    table.setCellTextDirection(0, 0, undefined);
    table.setCellTextDirection(0, 1, 'wordArtVert');
    const editedDirections = [
      undefined,
      'wordArtVert',
      undefined,
      'vert',
      'wordArtVert',
    ];
    expect(table.rows[0]!.cells.map(({ textDirection }) => textDirection))
      .toEqual(editedDirections);
    expect(duplicateTable!.rows[0]!.cells.map(
      ({ textDirection }) => textDirection)).toEqual(materializedDirections);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri);
    const reopenedTable = reopenedSlide?.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel
        && shape.name === 'Table text direction lifecycle',
    );
    const reopenedDuplicate = reopened.slides
      .find(({ partUri }) => partUri === duplicate.partUri)
      ?.shapes.find((shape): shape is TableModel => shape instanceof TableModel);
    expect(reopenedTable).toBeInstanceOf(TableModel);
    expect(reopenedDuplicate).toBeInstanceOf(TableModel);
    expect(reopenedTable!.rows[0]!.cells.map(
      ({ textDirection }) => textDirection)).toEqual(editedDirections);
    expect(reopenedDuplicate!.rows[0]!.cells.map(
      ({ textDirection }) => textDirection)).toEqual(materializedDirections);
    expect(reopenedTable!.columnWidths).toEqual(Array(5).fill(inches(2)));
    expect(reopenedTable!.rowHeights).toEqual([inches(1)]);
  });

  it('materializes table margin through edit, duplicate, write, and reopen', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const options: AddTableOptions = {
      name: 'Table margin lifecycle',
      margin: { top: 9, left: 18 },
      columnWidths: inches(2),
      rowHeights: inches(1),
    };
    const table = slide.addTable([[
      'Inherited string',
      { text: 'Partial override', options: { margin: { bottom: 12 } } },
      { text: 'Zero override', options: { margin: 0 } },
      { text: 'Tuple override', options: { margin: [1, 2, 3, 4] } },
    ]], options);
    const materializedMargins = [
      { top: 9, right: 7.2, bottom: 3.6, left: 18 },
      { top: 9, right: 7.2, bottom: 12, left: 18 },
      { top: 0, right: 0, bottom: 0, left: 0 },
      { top: 1, right: 2, bottom: 3, left: 4 },
    ];
    expect(table.rows[0]!.cells.map(({ margins }) => margins))
      .toEqual(materializedMargins);

    const sourceIndex = model.slides.indexOf(slide);
    const duplicate = model.duplicateSlide(sourceIndex);
    const duplicateTable = duplicate.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(duplicateTable).toBeInstanceOf(TableModel);
    expect(duplicateTable!.rows[0]!.cells.map(({ margins }) => margins))
      .toEqual(materializedMargins);

    table.setCellMargins(0, 0, undefined);
    table.setCellMargins(0, 1, { right: 5 });
    expect(table.rows[0]!.cells.map(({ margins }) => margins)).toEqual([
      undefined,
      { right: 5 },
      { top: 0, right: 0, bottom: 0, left: 0 },
      { top: 1, right: 2, bottom: 3, left: 4 },
    ]);
    expect(duplicateTable!.rows[0]!.cells.map(({ margins }) => margins))
      .toEqual(materializedMargins);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri);
    const reopenedTable = reopenedSlide?.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel
        && shape.name === 'Table margin lifecycle',
    );
    expect(reopenedTable).toBeInstanceOf(TableModel);
    expect(reopenedTable!.rows[0]!.cells.map(({ margins }) => margins)).toEqual([
      undefined,
      { right: 5 },
      { top: 0, right: 0, bottom: 0, left: 0 },
      { top: 1, right: 2, bottom: 3, left: 4 },
    ]);
    expect(reopenedTable!.columnWidths).toEqual(Array(4).fill(inches(2)));
    expect(reopenedTable!.rowHeights).toEqual([inches(1)]);

    const reopenedDuplicateSlide = reopened.slides.find(
      ({ partUri }) => partUri === duplicate.partUri,
    );
    const reopenedDuplicateTable = reopenedDuplicateSlide?.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel
        && shape.name === 'Table margin lifecycle',
    );
    expect(reopenedDuplicateTable).toBeInstanceOf(TableModel);
    expect(reopenedDuplicateTable!.rows[0]!.cells.map(({ margins }) => margins))
      .toEqual(materializedMargins);
  });

  it('materializes table fill through edit, duplicate, rollback, write, and reopen', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const sourceColor = {
      kind: 'scheme' as const,
      value: 'accent1' as 'accent1' | 'accent6',
    };
    const sourceFill = {
      kind: 'solid' as const,
      color: sourceColor,
      transparency: 33.3334,
    };
    const options: AddTableOptions = {
      name: 'Table fill lifecycle',
      fill: sourceFill,
      columnWidths: inches(2),
      rowHeights: inches(1),
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
    ]], options);
    const materializedFills = [
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
    ];
    expect(table.rows[0]!.cells.map(({ fill }) => fill)).toEqual(materializedFills);

    sourceColor.value = 'accent6';
    sourceFill.transparency = 1;
    expect(table.rows[0]!.cells.map(({ fill }) => fill)).toEqual(materializedFills);

    const sourceIndex = model.slides.indexOf(slide);
    const duplicate = model.duplicateSlide(sourceIndex);
    const duplicateTable = duplicate.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(duplicateTable).toBeInstanceOf(TableModel);
    expect(duplicateTable!.rows[0]!.cells.map(({ fill }) => fill))
      .toEqual(materializedFills);

    table.setCellFill(0, 0, undefined);
    table.setCellFill(0, 1, {
      kind: 'solid',
      color: { kind: 'srgb', value: '00FF00' },
      transparency: 0,
    });
    const editedFills = [
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
    ];
    expect(table.rows[0]!.cells.map(({ fill }) => fill)).toEqual(editedFills);
    expect(duplicateTable!.rows[0]!.cells.map(({ fill }) => fill))
      .toEqual(materializedFills);

    const beforeRollback = pkg.requirePart(slide.partUri).bytes.slice();
    const rollbackJournal = [...pkg.mutations];
    let rolledBack: TableModel | undefined;
    expect(() => pkg.transaction(() => {
      table.setCellFill(0, 2, {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF0000' },
      });
      rolledBack = slide.addTable(
        [['Temporary']],
        { fill: { kind: 'none' } },
      );
      throw new Error('restore table fill defaults');
    })).toThrow('restore table fill defaults');
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(slide.shapes).toHaveLength(1);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[0]!.cells.map(({ fill }) => fill)).toEqual(editedFills);
    expect(duplicateTable!.rows[0]!.cells.map(({ fill }) => fill))
      .toEqual(materializedFills);
    expect(() => rolledBack!.rows).toThrow(ModelParseError);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri);
    const reopenedTable = reopenedSlide?.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel
        && shape.name === 'Table fill lifecycle',
    );
    expect(reopenedTable).toBeInstanceOf(TableModel);
    expect(reopenedTable!.rows[0]!.cells.map(({ fill }) => fill)).toEqual(editedFills);
    expect(reopenedTable!.columnWidths).toEqual(Array(4).fill(inches(2)));
    expect(reopenedTable!.rowHeights).toEqual([inches(1)]);

    const reopenedDuplicateSlide = reopened.slides.find(
      ({ partUri }) => partUri === duplicate.partUri,
    );
    const reopenedDuplicateTable = reopenedDuplicateSlide?.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel
        && shape.name === 'Table fill lifecycle',
    );
    expect(reopenedDuplicateTable).toBeInstanceOf(TableModel);
    expect(reopenedDuplicateTable!.rows[0]!.cells.map(({ fill }) => fill))
      .toEqual(materializedFills);
    expect(reopenedDuplicateTable!.columnWidths).toEqual(Array(4).fill(inches(2)));
    expect(reopenedDuplicateTable!.rowHeights).toEqual([inches(1)]);
  });

  it('materializes table borders through edit, duplicate, rollback, write, and reopen', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
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
    const options: AddTableOptions = {
      name: 'Table border lifecycle',
      border: sourceBorder,
      columnWidths: inches(2),
      rowHeights: inches(1),
    };
    const table = slide.addTable([[
      'Inherited string',
      { text: 'Inherited object', options: {} },
      { text: 'Partial override', options: { border: {
        bottom: {
          kind: 'line',
          color: { kind: 'srgb', value: '70AD47' },
          width: 3,
          style: 'solid',
        },
      } } },
      { text: 'None override', options: { border: { kind: 'none' } } },
    ]], options);
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
    const materializedBorders = [
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
    expect(table.rows[0]!.cells.map(({ borders }) => borders))
      .toEqual(materializedBorders);

    sourceColor.value = 'accent6';
    sourceBorder.width = 9;
    expect(table.rows[0]!.cells.map(({ borders }) => borders))
      .toEqual(materializedBorders);

    const sourceIndex = model.slides.indexOf(slide);
    const duplicate = model.duplicateSlide(sourceIndex);
    const duplicateTable = duplicate.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel,
    );
    expect(duplicateTable).toBeInstanceOf(TableModel);
    expect(duplicateTable!.rows[0]!.cells.map(({ borders }) => borders))
      .toEqual(materializedBorders);

    table.setCellBorders(0, 0, undefined);
    table.setCellBorders(0, 1, {
      right: {
        kind: 'line',
        color: { kind: 'srgb', value: '00FF00' },
        width: 0,
        style: 'solid',
      },
    });
    const editedBorders = [
      undefined,
      {
        right: {
          kind: 'line',
          color: { kind: 'srgb', value: '00FF00' },
          width: 0,
          style: 'solid',
        },
      },
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
    expect(table.rows[0]!.cells.map(({ borders }) => borders)).toEqual(editedBorders);
    expect(duplicateTable!.rows[0]!.cells.map(({ borders }) => borders))
      .toEqual(materializedBorders);

    const beforeRollback = pkg.requirePart(slide.partUri).bytes.slice();
    const rollbackJournal = [...pkg.mutations];
    let rolledBack: TableModel | undefined;
    expect(() => pkg.transaction(() => {
      table.setCellBorders(0, 2, {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        width: 2,
      });
      rolledBack = slide.addTable(
        [['Temporary']],
        { border: { kind: 'none' } },
      );
      throw new Error('restore table border defaults');
    })).toThrow('restore table border defaults');
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(slide.shapes).toHaveLength(1);
    expect(slide.shapes[0]).toBe(table);
    expect(table.rows[0]!.cells.map(({ borders }) => borders)).toEqual(editedBorders);
    expect(duplicateTable!.rows[0]!.cells.map(({ borders }) => borders))
      .toEqual(materializedBorders);
    expect(() => rolledBack!.rows).toThrow(ModelParseError);

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    const reopenedSlide = reopened.slides.find(({ partUri }) => partUri === slide.partUri);
    const reopenedTable = reopenedSlide?.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel
        && shape.name === 'Table border lifecycle',
    );
    expect(reopenedTable).toBeInstanceOf(TableModel);
    expect(reopenedTable!.rows[0]!.cells.map(({ borders }) => borders))
      .toEqual(editedBorders);
    expect(reopenedTable!.columnWidths).toEqual(Array(4).fill(inches(2)));
    expect(reopenedTable!.rowHeights).toEqual([inches(1)]);

    const reopenedDuplicateSlide = reopened.slides.find(
      ({ partUri }) => partUri === duplicate.partUri,
    );
    const reopenedDuplicateTable = reopenedDuplicateSlide?.shapes.find(
      (shape): shape is TableModel => shape instanceof TableModel
        && shape.name === 'Table border lifecycle',
    );
    expect(reopenedDuplicateTable).toBeInstanceOf(TableModel);
    expect(reopenedDuplicateTable!.rows[0]!.cells.map(({ borders }) => borders))
      .toEqual(materializedBorders);
    expect(reopenedDuplicateTable!.columnWidths).toEqual(Array(4).fill(inches(2)));
    expect(reopenedDuplicateTable!.rowHeights).toEqual([inches(1)]);
  });

  it('reads and losslessly edits live table column widths through no-op, repair, merge, and rollback', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const table = slide.addTable(
      [
        ['A', 'B', 'C'],
        ['D', 'E', 'F'],
      ],
      {
        columnWidths: [inches(1), inches(2), inches(3)],
        height: inches(2),
      },
    );
    const part = pkg.requirePart(slide.partUri);
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes)
        .replace(
          '<a:tblGrid>',
          '<a:tblGrid keep="GRID">' +
            '<x:opaque xmlns:x="urn:test"><a:gridCol w="1"/></x:opaque>',
        )
        .replace('<a:tc>', '<a:tc gridSpan="2">')
        .replace('<a:tc>', '<a:tc hMerge="1" vMerge="1">'),
      part.contentType,
    );

    expect(table.columnWidths).toEqual([
      inches(1),
      inches(2),
      inches(3),
    ]);
    const detached = table.columnWidths as number[];
    detached[0] = 1;
    expect(table.columnWidths).toEqual([
      inches(1),
      inches(2),
      inches(3),
    ]);

    const beforeNoOp = pkg.requirePart(slide.partUri).bytes;
    const noOpJournal = [...pkg.mutations];
    table.setColumnWidths([inches(1), inches(2), inches(3)]);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeNoOp);
    expect(pkg.mutations).toEqual(noOpJournal);

    table.setColumnWidths([inches(1.5), inches(2.5), inches(2)]);
    expect(table.columnWidths).toEqual([
      inches(1.5),
      inches(2.5),
      inches(2),
    ]);
    expect(table.transform.width).toBe(inches(6));
    expect(slide.shapes.find(({ id }) => id === table.id)).toBe(table);
    let updated = new TextDecoder().decode(
      pkg.requirePart(slide.partUri).bytes,
    );
    expect(updated).toContain(
      '<a:tblGrid keep="GRID">' +
        '<x:opaque xmlns:x="urn:test"><a:gridCol w="1"/></x:opaque>',
    );
    expect(updated).toContain('gridSpan="2"');
    expect(updated).toContain('hMerge="1" vMerge="1"');

    const mismatchedPart = pkg.requirePart(slide.partUri);
    pkg.setPart(
      mismatchedPart.uri,
      new TextDecoder().decode(mismatchedPart.bytes)
        .replace('cx="5486400"', 'cx="914400"'),
      mismatchedPart.contentType,
    );
    expect(table.columnWidths).toEqual([
      inches(1.5),
      inches(2.5),
      inches(2),
    ]);
    expect(table.transform.width).toBe(inches(1));
    table.setColumnWidths(table.columnWidths!);
    expect(table.transform.width).toBe(inches(6));
    updated = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(updated).toContain('<a:ext cx="5486400" cy="1828800"/>');

    const beforeRollback = pkg.requirePart(slide.partUri).bytes;
    const rollbackJournal = [...pkg.mutations];
    expect(() =>
      pkg.transaction(() => {
        table.setColumnWidths(inches(1));
        expect(table.columnWidths).toEqual([
          inches(1),
          inches(1),
          inches(1),
        ]);
        throw new Error('restore table column widths');
      }),
    ).toThrow('restore table column widths');
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(table.columnWidths).toEqual([
      inches(1.5),
      inches(2.5),
      inches(2),
    ]);
    expect(table.transform.width).toBe(inches(6));
    expect(slide.shapes.find(({ id }) => id === table.id)).toBe(table);
  });

  it('rejects invalid table column-width inputs and malformed OOXML without mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const table = slide.addTable(
      [['A', 'B', 'C']],
      { columnWidths: [inches(1), inches(2), inches(3)] },
    );
    const accessor = [1, 2, 3];
    let accessorCalls = 0;
    Object.defineProperty(accessor, '1', {
      get() {
        accessorCalls += 1;
        return 2;
      },
      enumerable: true,
      configurable: true,
    });
    const hole = [1, 2, 3];
    delete hole[1];
    const extra = [1, 2, 3];
    Object.defineProperty(extra, 'extra', { value: true });
    const symbol = [1, 2, 3];
    Object.defineProperty(symbol, Symbol('width'), { value: true });
    const invalid = [
      undefined,
      null,
      [],
      [1],
      [1, 2],
      [1, 2, 3, 4],
      new Uint32Array([1, 2, 3]),
      Number.NaN,
      Number.POSITIVE_INFINITY,
      0,
      -1,
      Number.MAX_SAFE_INTEGER,
      accessor,
      hole,
      extra,
      symbol,
      [Number.MAX_SAFE_INTEGER, 1, 1],
    ];
    const beforeInvalid = pkg.requirePart(slide.partUri).bytes;
    const invalidJournal = [...pkg.mutations];
    for (const value of invalid) {
      expect(() => table.setColumnWidths(value as never)).toThrow();
    }
    expect(accessorCalls).toBe(0);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);

    const malformedGridPart = pkg.requirePart(slide.partUri);
    pkg.setPart(
      malformedGridPart.uri,
      new TextDecoder().decode(malformedGridPart.bytes)
        .replace('w="914400"', 'w="0"'),
      malformedGridPart.contentType,
    );
    expect(table.columnWidths).toBeUndefined();
    const beforeGridFailure = pkg.requirePart(slide.partUri).bytes;
    const gridFailureJournal = [...pkg.mutations];
    expect(() => table.setColumnWidths(inches(1))).toThrow(ModelParseError);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeGridFailure);
    expect(pkg.mutations).toEqual(gridFailureJournal);

    const transformPkg = await OpcPackage.open(await modelFixture());
    const transformModel = new PresentationModel(transformPkg);
    const transformSlide = transformModel.addSlide();
    const transformTable = transformSlide.addTable(
      [['A', 'B']],
      { columnWidths: [inches(1), inches(2)] },
    );
    const transformPart = transformPkg.requirePart(transformSlide.partUri);
    transformPkg.setPart(
      transformPart.uri,
      new TextDecoder().decode(transformPart.bytes)
        .replace('cx="2743200"', 'x:cx="2743200" xmlns:x="urn:test"'),
      transformPart.contentType,
    );
    expect(transformTable.columnWidths).toEqual([inches(1), inches(2)]);
    const beforeTransformFailure =
      transformPkg.requirePart(transformSlide.partUri).bytes;
    const transformFailureJournal = [...transformPkg.mutations];
    let thrown: unknown;
    try {
      transformTable.setColumnWidths(inches(1));
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(ModelParseError);
    expect(thrown).toMatchObject({ partUri: transformSlide.partUri });
    expect(transformPkg.requirePart(transformSlide.partUri).bytes)
      .toEqual(beforeTransformFailure);
    expect(transformPkg.mutations).toEqual(transformFailureJournal);
  });

  it('reads and losslessly edits live table row heights through explicit, automatic, merge, no-op, repair, and rollback paths', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const table = slide.addTable(
      [
        ['A', 'B'],
        ['C', 'D'],
        ['E', 'F'],
      ],
      {
        columnWidths: [inches(2), inches(3)],
        rowHeights: [inches(0.5), inches(0.75), inches(1)],
      },
    );
    const part = pkg.requirePart(slide.partUri);
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes)
        .replace(
          '</a:tblGrid>',
          '</a:tblGrid>' +
            '<x:opaque xmlns:x="urn:test"><a:tr h="999"/></x:opaque>',
        )
        .replace('<a:tc>', '<a:tc rowSpan="2">')
        .replace('<a:tc>', '<a:tc hMerge="1" vMerge="1">')
        .replace('h="457200"', 'h="0457200"')
        .replace('cy="2057400"', 'cy="02057400"'),
      part.contentType,
    );

    expect(table.rowHeights).toEqual([
      inches(0.5),
      inches(0.75),
      inches(1),
    ]);
    const detached = table.rowHeights as number[];
    detached[0] = 1;
    expect(table.rowHeights).toEqual([
      inches(0.5),
      inches(0.75),
      inches(1),
    ]);

    const beforeNoOp = pkg.requirePart(slide.partUri).bytes;
    const noOpJournal = [...pkg.mutations];
    table.setRowHeights([inches(0.5), inches(0.75), inches(1)]);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeNoOp);
    expect(pkg.mutations).toEqual(noOpJournal);

    table.setRowHeights([inches(0.75), inches(1.25), inches(0.5)]);
    expect(table.rowHeights).toEqual([
      inches(0.75),
      inches(1.25),
      inches(0.5),
    ]);
    expect(table.transform.height).toBe(inches(2.5));
    expect(slide.shapes.find(({ id }) => id === table.id)).toBe(table);
    let updated = new TextDecoder().decode(
      pkg.requirePart(slide.partUri).bytes,
    );
    expect(updated).toContain(
      '<x:opaque xmlns:x="urn:test"><a:tr h="999"/></x:opaque>',
    );
    expect(updated).toContain('rowSpan="2"');
    expect(updated).toContain('hMerge="1" vMerge="1"');

    table.setTransform({ height: inches(3) });
    expect(table.transform.height).toBe(inches(3));
    expect(table.rowHeights).toEqual([
      inches(0.75),
      inches(1.25),
      inches(0.5),
    ]);
    table.setRowHeights(table.rowHeights!);
    expect(table.transform.height).toBe(inches(2.5));

    const mixedTransformHeight = table.transform.height;
    table.setRowHeights([0, inches(1), 0]);
    expect(table.rowHeights).toEqual([0, inches(1), 0]);
    expect(table.transform.height).toBe(mixedTransformHeight);
    updated = new TextDecoder().decode(pkg.requirePart(slide.partUri).bytes);
    expect(updated).toContain('<a:ext cx="4572000" cy="2286000"/>');

    const beforeMixedNoOp = pkg.requirePart(slide.partUri).bytes;
    const mixedNoOpJournal = [...pkg.mutations];
    table.setRowHeights([0, inches(1), 0]);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeMixedNoOp);
    expect(pkg.mutations).toEqual(mixedNoOpJournal);

    const beforeRollback = pkg.requirePart(slide.partUri).bytes;
    const rollbackJournal = [...pkg.mutations];
    expect(() =>
      pkg.transaction(() => {
        table.setRowHeights(inches(1));
        expect(table.rowHeights).toEqual([
          inches(1),
          inches(1),
          inches(1),
        ]);
        expect(table.transform.height).toBe(inches(3));
        throw new Error('restore table row heights');
      }),
    ).toThrow('restore table row heights');
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(table.rowHeights).toEqual([0, inches(1), 0]);
    expect(table.transform.height).toBe(mixedTransformHeight);
    expect(slide.shapes.find(({ id }) => id === table.id)).toBe(table);
  });

  it('rejects invalid table row-height inputs and malformed OOXML without mutation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.addSlide();
    const table = slide.addTable(
      [['A'], ['B'], ['C']],
      { rowHeights: [inches(0.5), inches(0.75), inches(1)] },
    );
    const accessor = [1, 2, 3];
    let accessorCalls = 0;
    Object.defineProperty(accessor, '1', {
      get() {
        accessorCalls += 1;
        return 2;
      },
      enumerable: true,
      configurable: true,
    });
    const hole = [1, 2, 3];
    delete hole[1];
    const extra = [1, 2, 3];
    Object.defineProperty(extra, 'extra', { value: true });
    const symbol = [1, 2, 3];
    Object.defineProperty(symbol, Symbol('height'), { value: true });
    const invalid = [
      undefined,
      null,
      [],
      [1],
      [1, 2],
      [1, 2, 3, 4],
      new Uint32Array([1, 2, 3]),
      Number.NaN,
      Number.POSITIVE_INFINITY,
      -1,
      -0.4,
      Number.MAX_SAFE_INTEGER,
      accessor,
      hole,
      extra,
      symbol,
      [Number.MAX_SAFE_INTEGER, 1, 1],
    ];
    const beforeInvalid = pkg.requirePart(slide.partUri).bytes;
    const invalidJournal = [...pkg.mutations];
    for (const value of invalid) {
      expect(() => table.setRowHeights(value as never)).toThrow();
    }
    expect(accessorCalls).toBe(0);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);

    const malformedRowPart = pkg.requirePart(slide.partUri);
    pkg.setPart(
      malformedRowPart.uri,
      new TextDecoder().decode(malformedRowPart.bytes)
        .replace('h="457200"', 'x:h="457200" xmlns:x="urn:test"'),
      malformedRowPart.contentType,
    );
    expect(table.rowHeights).toBeUndefined();
    const beforeRowFailure = pkg.requirePart(slide.partUri).bytes;
    const rowFailureJournal = [...pkg.mutations];
    expect(() => table.setRowHeights(inches(1))).toThrow(ModelParseError);
    expect(pkg.requirePart(slide.partUri).bytes).toEqual(beforeRowFailure);
    expect(pkg.mutations).toEqual(rowFailureJournal);

    const transformPkg = await OpcPackage.open(await modelFixture());
    const transformModel = new PresentationModel(transformPkg);
    const transformSlide = transformModel.addSlide();
    const transformTable = transformSlide.addTable(
      [['A'], ['B']],
      { rowHeights: [inches(0.5), inches(1.5)] },
    );
    const transformPart = transformPkg.requirePart(transformSlide.partUri);
    transformPkg.setPart(
      transformPart.uri,
      new TextDecoder().decode(transformPart.bytes)
        .replace('cy="1828800"', 'x:cy="1828800" xmlns:x="urn:test"'),
      transformPart.contentType,
    );
    expect(transformTable.rowHeights).toEqual([
      inches(0.5),
      inches(1.5),
    ]);
    const beforeTransformFailure =
      transformPkg.requirePart(transformSlide.partUri).bytes;
    const transformFailureJournal = [...transformPkg.mutations];
    let thrown: unknown;
    try {
      transformTable.setRowHeights(0);
    } catch (error) {
      thrown = error;
    }
    expect(thrown).toBeInstanceOf(ModelParseError);
    expect(thrown).toMatchObject({ partUri: transformSlide.partUri });
    expect(transformPkg.requirePart(transformSlide.partUri).bytes)
      .toEqual(beforeTransformFailure);
    expect(transformPkg.mutations).toEqual(transformFailureJournal);
  });

  it('reads only exact direct table-cell text directions into detached snapshots', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const part = pkg.requirePart('/ppt/slides/slide1.xml');
    const cell = (properties: string, index: number, bodyProperties = '<a:bodyPr/>'): string =>
      `<a:tc><a:txBody>${bodyProperties}<a:p><a:r><a:t>Cell ${index}</a:t></a:r></a:p></a:txBody>${properties}</a:tc>`;
    const cells = [
      cell('<a:tcPr vert="horz"/>', 0),
      cell('<a:tcPr vert="vert"/>', 1),
      cell('<a:tcPr vert="vert270"/>', 2),
      cell('<a:tcPr vert="wordArtVert"/>', 3),
      cell('<a:tcPr keep="absent"/>', 4),
      cell('<a:tcPr vert=""/>', 5),
      cell('<a:tcPr vert="Vert"/>', 6),
      cell('<a:tcPr vert=" vert "/>', 7),
      cell('<a:tcPr xmlns:x="urn:test" x:vert="vert"/>', 8),
      cell('<a:tcPr vert="vert" vert="horz"/>', 9),
      cell('<a:tcPr vert="eaVert"/>', 10),
      cell('<a:tcPr keep="body-direction"/>', 11, '<a:bodyPr vert="vert" custom="KEEP"/>'),
      cell('<a:tcPr vert="vert"/><a:tcPr keep="repeated"/>', 12, '<a:bodyPr custom="KEEP"/>'),
      cell('', 13, '<a:bodyPr custom="KEEP"/>'),
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        /<a:tr>.*?<\/a:tr>/,
        `<a:tr>${cells}</a:tr>`,
      ),
      part.contentType,
    );
    const table = new PresentationModel(pkg).slides[1]!.shapes[2] as TableModel;
    const journal = [...pkg.mutations];

    const snapshot = table.rows;
    const directions = snapshot[0]!.cells.map(({ textDirection }) => textDirection);
    expect(directions.slice(0, 4)).toEqual(['horz', 'vert', 'vert270', 'wordArtVert']);
    expect(directions.slice(4)).toEqual(Array(10).fill(undefined));
    expect(pkg.mutations).toEqual(journal);

    (snapshot[0]!.cells[0] as { textDirection?: string }).textDirection = 'vert';
    expect(table.rows[0]!.cells[0]!.textDirection).toBe('horz');
    expect(pkg.mutations).toEqual(journal);
  });

  it('losslessly edits one direct table-cell direction and rolls back atomically', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const part = pkg.requirePart('/ppt/slides/slide1.xml');
    const cell = (properties: string, text: string, cellAttributes = ''): string =>
      `<a:tc${cellAttributes}><a:txBody><a:bodyPr custom="BODY"/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></a:txBody>${properties}</a:tc>`;
    const adjacentCell = cell('<a:tcPr vert="vert" keep="ADJACENT"/>', 'Adjacent');
    const cells = [
      cell("<a:tcPr vert='eaVert' marL=\"100\"><a:solidFill><a:srgbClr val=\"112233\"/></a:solidFill><x:keep xmlns:x=\"urn:test\"/></a:tcPr>", 'Replace'),
      cell('<a:tcPr marR="200"/>', 'Merged', ' hMerge="1"'),
      cell('<a:tcPr marT="300"><a:noFill/></a:tcPr>', 'Add'),
      adjacentCell,
      cell('<a:tcPr vert="vert"/><a:tcPr keep="AMBIGUOUS"/>', 'Repeated'),
      cell('', 'Missing'),
    ].join('');
    const source = new TextDecoder().decode(part.bytes)
      .replace(
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr><a:graphic>',
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="2000"/></p:xfrm><a:graphic>',
      )
      .replace(/<a:tr>.*?<\/a:tr>/, `<a:tr>${cells}</a:tr>`);
    pkg.setPart(part.uri, source, part.contentType);
    const model = new PresentationModel(pkg);
    const table = model.slides[1]!.shapes[2] as TableModel;

    const absentClearJournal = [...pkg.mutations];
    table.setCellTextDirection(0, 1, undefined);
    expect(pkg.mutations).toEqual(absentClearJournal);

    table.setCellTextDirection(0, 0, 'vert270');
    let updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain("<a:tcPr vert='vert270' marL=\"100\"><a:solidFill>");
    expect(updated).toContain(adjacentCell);
    expect(table.rows[0]!.cells[0]!.textDirection).toBe('vert270');

    table.setCellTextDirection(0, 1, 'wordArtVert');
    table.setCellTextDirection(0, 2, 'horz');
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain('<a:tcPr marR="200" vert="wordArtVert"/>');
    expect(updated).toContain('<a:tcPr marT="300" vert="horz"><a:noFill/></a:tcPr>');
    expect(updated).toContain('<a:tc hMerge="1">');
    expect(updated).toContain('<a:bodyPr custom="BODY"/>');
    expect(updated).toContain(adjacentCell);

    const noOpJournal = [...pkg.mutations];
    table.setCellTextDirection(0, 2, 'horz');
    expect(pkg.mutations).toEqual(noOpJournal);

    table.setCellTextDirection(0, 0, undefined);
    table.setCellText(0, 2, 'Edited Add');
    table.setTransform({ x: inches(1) });
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain('<a:tcPr marL="100"><a:solidFill>');
    expect(updated).not.toContain("vert='vert270'");
    expect(updated).toContain('<a:tcPr marT="300" vert="horz"><a:noFill/></a:tcPr>');
    expect(updated).toContain('<a:t>Edited Add</a:t>');
    expect(updated).toContain('<a:off x="914400" y="0"/>');
    expect(updated).toContain(adjacentCell);

    const beforeInvalid = pkg.requirePart(part.uri).bytes;
    const invalidJournal = [...pkg.mutations];
    for (const [row, column] of [[-1, 0], [0, -1], [0, 9], [1, 0]]) {
      expect(() => table.setCellTextDirection(row!, column!, 'vert')).toThrow(RangeError);
    }
    expect(() => table.setCellTextDirection(0, 4, 'vert270')).toThrow(ModelParseError);
    expect(() => table.setCellTextDirection(0, 5, 'vert270')).toThrow(ModelParseError);
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);

    const beforeRollback = pkg.requirePart(part.uri).bytes;
    const rollbackJournal = [...pkg.mutations];
    const rollbackDirections = table.rows[0]!.cells.map(({ textDirection }) => textDirection);
    expect(() =>
      pkg.transaction(() => {
        table.setCellTextDirection(0, 3, 'wordArtVert');
        table.setCellTextDirection(0, 1, undefined);
        throw new Error('restore table cell directions');
      }),
    ).toThrow('restore table cell directions');
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(model.slides[1]!.shapes[2]).toBe(table);
    expect(table.rows[0]!.cells.map(({ textDirection }) => textDirection)).toEqual(rollbackDirections);
  });

  it('reads only a unique direct table-cell fit choice into detached snapshots', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const part = pkg.requirePart('/ppt/slides/slide1.xml');
    const textBody = (bodyProperties: string, index: number): string =>
      `<a:txBody>${bodyProperties}<a:p><a:r><a:t>Fit ${index}</a:t></a:r></a:p></a:txBody>`;
    const cell = (body: string, index: number): string =>
      `<a:tc>${body}<a:tcPr vert="horz" keep="CELL-${index}"/></a:tc>`;
    const cells = [
      cell(textBody('<a:bodyPr><a:noAutofit/></a:bodyPr>', 0), 0),
      cell(textBody("<a:bodyPr><a:normAutofit fontScale='85000' lnSpcReduction=\"20000\"/></a:bodyPr>", 1), 1),
      cell(textBody('<a:bodyPr><a:spAutoFit/></a:bodyPr>', 2), 2),
      cell(textBody('<a:bodyPr/>', 3), 3),
      cell(textBody('<a:bodyPr><a:normAutofit/><a:normAutofit/></a:bodyPr>', 4), 4),
      cell(textBody('<a:bodyPr><a:noAutofit/><a:spAutoFit/></a:bodyPr>', 5), 5),
      cell(textBody('<a:bodyPr><a:normAutoFit/></a:bodyPr>', 6), 6),
      cell(textBody('<a:bodyPr xmlns:x="urn:test"><x:normAutofit/></a:bodyPr>', 7), 7),
      cell(textBody('<a:bodyPr><a:extLst><a:normAutofit/></a:extLst></a:bodyPr>', 8), 8),
      cell(
        `${textBody('<a:bodyPr><a:noAutofit/></a:bodyPr>', 9)}${textBody('<a:bodyPr custom="SECOND"/>', 9)}`,
        9,
      ),
      cell(
        textBody('<a:bodyPr><a:spAutoFit/></a:bodyPr><a:bodyPr custom="SECOND"/>', 10),
        10,
      ),
      cell('<a:txBody><a:p><a:r><a:t>Fit 11</a:t></a:r></a:p></a:txBody>', 11),
      cell('<x:keep xmlns:x="urn:test">NO TEXT BODY</x:keep>', 12),
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(/<a:tr>.*?<\/a:tr>/, `<a:tr>${cells}</a:tr>`),
      part.contentType,
    );
    const table = new PresentationModel(pkg).slides[1]!.shapes[2] as TableModel;
    const journal = [...pkg.mutations];

    const snapshot = table.rows;
    const fits = snapshot[0]!.cells.map(({ textFit }) => textFit);
    expect(fits.slice(0, 3)).toEqual(['none', 'shrink', 'resize']);
    expect(fits.slice(3)).toEqual(Array(10).fill(undefined));
    expect(snapshot[0]!.cells.every(({ textDirection }) => textDirection === 'horz')).toBe(true);
    expect(pkg.mutations).toEqual(journal);

    (snapshot[0]!.cells[0] as { textFit?: string }).textFit = 'resize';
    expect(table.rows[0]!.cells[0]!.textFit).toBe('none');
    expect(pkg.mutations).toEqual(journal);
  });

  it('losslessly edits one table-cell fit choice and rolls back atomically', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const part = pkg.requirePart('/ppt/slides/slide1.xml');
    const textBody = (bodyProperties: string, text: string): string =>
      `<a:txBody>${bodyProperties}<a:p><a:r><a:t>${text}</a:t></a:r></a:p></a:txBody>`;
    const cell = (body: string, properties: string, attributes = ''): string =>
      `<a:tc${attributes}>${body}${properties}</a:tc>`;
    const adjacentCell = cell(
      textBody('<a:bodyPr keep="ADJACENT"><a:spAutoFit/></a:bodyPr>', 'Adjacent'),
      '<a:tcPr vert="horz" keep="ADJACENT-TCPR"/>',
    );
    const cells = [
      cell(
        textBody(
          '<a:bodyPr wrap="none" lIns="127000" anchor="b" vert="vert270" custom="KEEP"><a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp><a:normAutofit fontScale=\'85000\' lnSpcReduction="20000"/><a:scene3d><a:camera prst="orthographicFront"/></a:scene3d><x:keep xmlns:x="urn:test">KEEP</x:keep></a:bodyPr>',
          'Calculated',
        ),
        '<a:tcPr vert="vert270" keep="TCPR"/>',
      ),
      cell(textBody('<a:bodyPr/>', 'Self closing'), '<a:tcPr vert="vert" keep="SELF"/>'),
      cell(
        textBody(
          '<a:bodyPr custom="ORDER"><a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp><a:scene3d><a:camera prst="orthographicFront"/></a:scene3d><a:extLst><a:ext uri="urn:test"/></a:extLst></a:bodyPr>',
          'Ordered',
        ),
        '<a:tcPr vert="horz" keep="ORDER"/>',
      ),
      cell(
        textBody('<a:bodyPr><a:noAutofit/><a:spAutoFit/><x:keep xmlns:x="urn:test">CONFLICT</x:keep></a:bodyPr>', 'Conflict'),
        '<a:tcPr keep="CONFLICT"/>',
      ),
      adjacentCell,
      cell(
        textBody('<a:bodyPr custom="MERGED"><a:noAutofit/></a:bodyPr>', 'Merged'),
        '<a:tcPr vert="wordArtVert" keep="MERGED"/>',
        ' hMerge="1"',
      ),
      cell('<a:txBody><a:p><a:r><a:t>Missing bodyPr</a:t></a:r></a:p></a:txBody>', '<a:tcPr/>'),
      cell(
        textBody('<a:bodyPr/><a:bodyPr custom="SECOND"/>', 'Repeated bodyPr'),
        '<a:tcPr/>',
      ),
      cell(
        `${textBody('<a:bodyPr/>', 'First txBody')}${textBody('<a:bodyPr custom="SECOND"/>', 'Second txBody')}`,
        '<a:tcPr/>',
      ),
      cell('<x:keep xmlns:x="urn:test">MISSING TXBODY</x:keep>', '<a:tcPr/>'),
    ].join('');
    const source = new TextDecoder().decode(part.bytes)
      .replace(
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr><a:graphic>',
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="2000"/></p:xfrm><a:graphic>',
      )
      .replace(/<a:tr>.*?<\/a:tr>/, `<a:tr>${cells}</a:tr>`);
    pkg.setPart(part.uri, source, part.contentType);
    const model = new PresentationModel(pkg);
    const table = model.slides[1]!.shapes[2] as TableModel;

    const noOpJournal = [...pkg.mutations];
    table.setCellTextFit(0, 0, 'shrink');
    table.setCellTextFit(0, 1, undefined);
    expect(pkg.mutations).toEqual(noOpJournal);

    table.setCellTextFit(0, 0, 'resize');
    let updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain('<a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp><a:spAutoFit/><a:scene3d>');
    expect(updated).not.toContain('fontScale=\'85000\'');
    expect(updated).toContain('<a:bodyPr wrap="none" lIns="127000" anchor="b" vert="vert270" custom="KEEP">');
    expect(updated).toContain('<a:tcPr vert="vert270" keep="TCPR"/>');
    expect(updated).toContain(adjacentCell);

    table.setCellTextFit(0, 1, 'shrink');
    table.setCellTextFit(0, 2, 'resize');
    table.setCellTextFit(0, 3, 'shrink');
    table.setCellTextFit(0, 5, 'none');
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain('<a:bodyPr><a:normAutofit/></a:bodyPr><a:p><a:r><a:t>Self closing</a:t>');
    expect(updated).toContain('<a:bodyPr custom="ORDER"><a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp><a:spAutoFit/><a:scene3d>');
    expect(updated).toContain('<a:bodyPr><a:normAutofit/><x:keep xmlns:x="urn:test">CONFLICT</x:keep></a:bodyPr>');
    expect(updated).toContain('<a:bodyPr custom="MERGED"></a:bodyPr>');
    expect(updated).toContain('<a:tc hMerge="1">');
    expect(updated).toContain(adjacentCell);

    table.setCellText(0, 2, 'Edited ordered');
    table.setCellTextDirection(0, 1, 'wordArtVert');
    table.setTransform({ x: inches(1) });
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain('<a:t>Edited ordered</a:t>');
    expect(updated).toContain('<a:tcPr vert="wordArtVert" keep="SELF"/>');
    expect(updated).toContain('<a:bodyPr><a:normAutofit/></a:bodyPr>');
    expect(updated).toContain('<a:bodyPr custom="ORDER"><a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp><a:spAutoFit/>');
    expect(updated).toContain('<a:off x="914400" y="0"/>');
    expect(updated).toContain(adjacentCell);

    table.setCellTextFit(0, 1, undefined);
    expect(table.rows[0]!.cells[1]!.textFit).toBeUndefined();
    expect(table.rows[0]!.cells[1]!.textDirection).toBe('wordArtVert');

    const beforeInvalid = pkg.requirePart(part.uri).bytes;
    const invalidJournal = [...pkg.mutations];
    for (const [row, column] of [[-1, 0], [0, -1], [0, 10], [1, 0]]) {
      expect(() => table.setCellTextFit(row!, column!, 'resize')).toThrow(RangeError);
    }
    for (const column of [6, 7, 8, 9]) {
      expect(() => table.setCellTextFit(0, column, 'resize')).toThrow(ModelParseError);
    }
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);

    const beforeRollback = pkg.requirePart(part.uri).bytes;
    const rollbackJournal = [...pkg.mutations];
    const rollbackFits = table.rows[0]!.cells.map(({ textFit }) => textFit);
    expect(() =>
      pkg.transaction(() => {
        table.setCellTextFit(0, 0, 'shrink');
        table.setCellTextFit(0, 2, undefined);
        throw new Error('restore table cell fits');
      }),
    ).toThrow('restore table cell fits');
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(model.slides[1]!.shapes[2]).toBe(table);
    expect(table.rows[0]!.cells.map(({ textFit }) => textFit)).toEqual(rollbackFits);
  });

  it('reads strict direct table-cell margins into detached point snapshots', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const part = pkg.requirePart('/ppt/slides/slide1.xml');
    const cell = (
      properties: string,
      index: number,
      bodyProperties = '<a:bodyPr/>',
    ): string =>
      `<a:tc><a:txBody>${bodyProperties}<a:p><a:r><a:t>Margins ${index}</a:t></a:r></a:p></a:txBody>${properties}</a:tc>`;
    const cells = [
      cell('<a:tcPr marT="45720" marR="91440" marB="137160" marL="182880"/>', 0),
      cell('<a:tcPr marT="0" marL="-12700"/>', 1),
      cell('<a:tcPr marR="1588"/>', 2),
      cell('<a:tcPr marT="-2147483648" marB="2147483647"/>', 3),
      cell('<a:tcPr vert="horz"/>', 4),
      cell('<a:tcPr marT="" marR="101600" marB="1.5" marL="1e3"/>', 5),
      cell(
        '<a:tcPr xmlns:x="urn:test" x:marL="12700"/>',
        6,
        '<a:bodyPr lIns="12700" tIns="25400" rIns="38100" bIns="50800"/>',
      ),
      cell('<a:tcPr marT="12700" marT="25400"/>', 7),
      cell('<a:tcPr mart="12700" marR=" 12700 "/>', 8),
      cell('<a:tcPr marL="-2147483649" marR="2147483648"/>', 9),
      cell('<a:tcPr marL="12700"/><a:tcPr marR="25400"/>', 10),
      cell('', 11),
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(/<a:tr>.*?<\/a:tr>/, `<a:tr>${cells}</a:tr>`),
      part.contentType,
    );
    const table = new PresentationModel(pkg).slides[1]!.shapes[2] as TableModel;
    const journal = [...pkg.mutations];

    const snapshot = table.rows;
    expect(snapshot[0]!.cells.map(({ margins }) => margins)).toEqual([
      { top: 3.6, right: 7.2, bottom: 10.8, left: 14.4 },
      { top: 0, left: -1 },
      { right: 1_588 / 12_700 },
      { top: -2_147_483_648 / 12_700, bottom: 2_147_483_647 / 12_700 },
      undefined,
      { right: 8 },
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    expect(pkg.mutations).toEqual(journal);

    (snapshot[0]!.cells[0]!.margins as { top?: number }).top = 99;
    expect(table.rows[0]!.cells[0]!.margins).toEqual({
      top: 3.6,
      right: 7.2,
      bottom: 10.8,
      left: 14.4,
    });
    expect(pkg.mutations).toEqual(journal);
  });

  it('losslessly replaces one table-cell margin set and rolls back atomically', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const part = pkg.requirePart('/ppt/slides/slide1.xml');
    const textBody = (bodyProperties: string, text: string): string =>
      `<a:txBody>${bodyProperties}<a:p><a:r><a:t>${text}</a:t></a:r></a:p></a:txBody>`;
    const cell = (body: string, properties: string, attributes = ''): string =>
      `<a:tc${attributes}>${body}${properties}</a:tc>`;
    const adjacentCell = cell(
      textBody('<a:bodyPr lIns="12700" keep="ADJACENT"><a:spAutoFit/></a:bodyPr>', 'Adjacent'),
      '<a:tcPr marL="12700" marR="25400" marT="38100" marB="50800" anchor="b" vert="horz" keep="ADJACENT-TCPR"/>',
    );
    const cells = [
      cell(
        textBody(
          '<a:bodyPr lIns="99999" tIns="88888" anchor="b" custom="BODY"><a:normAutofit fontScale="85000"/><x:keep xmlns:x="urn:test">BODY</x:keep></a:bodyPr>',
          'Same numeric margins',
        ),
        '<a:tcPr marL="00101600" marR="101600" marT="101600" marB="101600" anchor=\'ctr\' vert="vert270" custom="KEEP"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><x:keep xmlns:x="urn:test">TCPR</x:keep></a:tcPr>',
      ),
      cell(textBody('<a:bodyPr/>', 'Add uniform'), '<a:tcPr anchor="t"/>'),
      cell(
        textBody('<a:bodyPr><a:noAutofit/></a:bodyPr>', 'Replace TRBL'),
        '<a:tcPr marB=\'0\' marT="0" marR="0" marL="0" horzOverflow="clip"><a:noFill/></a:tcPr>',
      ),
      cell(
        textBody('<a:bodyPr/>', 'Partial object'),
        '<a:tcPr marT="1.5" marR="invalid" marB="25400" keep="PARTIAL"/>',
      ),
      cell(
        textBody('<a:bodyPr/>', 'Clear object'),
        '<a:tcPr marL="12700" marR="25400" marT="38100" marB="50800" anchor="b" keep="CLEAR-OBJECT"/>',
      ),
      cell(
        textBody('<a:bodyPr/>', 'Clear undefined'),
        '<a:tcPr xmlns:x="urn:test" marL="unknown" x:marL="777" vert="wordArtVert" keep="CLEAR-UNDEFINED"/>',
      ),
      adjacentCell,
      cell(
        textBody('<a:bodyPr><a:spAutoFit/></a:bodyPr>', 'Merged'),
        '<a:tcPr marL="12700" marR="12700" marT="12700" marB="12700" anchor="ctr" vert="wordArtVert" keep="MERGED"/>',
        ' hMerge="1"',
      ),
      cell(textBody('<a:bodyPr/>', 'Repeated margin'), '<a:tcPr marT="12700" marT="25400"/>'),
      cell(textBody('<a:bodyPr/>', 'Repeated tcPr'), '<a:tcPr/><a:tcPr keep="SECOND"/>'),
      cell(textBody('<a:bodyPr/>', 'Missing tcPr'), ''),
    ].join('');
    const source = new TextDecoder().decode(part.bytes)
      .replace(
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr><a:graphic>',
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="2000"/></p:xfrm><a:graphic>',
      )
      .replace(/<a:tr>.*?<\/a:tr>/, `<a:tr>${cells}</a:tr>`);
    pkg.setPart(part.uri, source, part.contentType);
    const model = new PresentationModel(pkg);
    const table = model.slides[1]!.shapes[2] as TableModel;

    const noOpJournal = [...pkg.mutations];
    table.setCellMargins(0, 0, 8);
    expect(pkg.mutations).toEqual(noOpJournal);
    expect(new TextDecoder().decode(pkg.requirePart(part.uri).bytes)).toContain(
      '<a:tcPr marL="00101600" marR="101600" marT="101600" marB="101600"',
    );

    table.setCellMargins(0, 1, 4);
    table.setCellMargins(0, 2, [1, 2, 3, 4]);
    table.setCellMargins(0, 3, { top: 4, left: 8 });
    table.setCellMargins(0, 4, {});
    table.setCellMargins(0, 5, undefined);
    table.setCellMargins(0, 7, { bottom: 6 });
    let updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain(
      '<a:tcPr anchor="t" marL="50800" marR="50800" marT="50800" marB="50800"/>',
    );
    expect(updated).toContain(
      '<a:tcPr marB=\'38100\' marT="12700" marR="25400" marL="50800" horzOverflow="clip">',
    );
    expect(updated).toContain('<a:tcPr marT="50800" keep="PARTIAL" marL="101600"/>');
    expect(updated).toContain('<a:tcPr anchor="b" keep="CLEAR-OBJECT"/>');
    expect(updated).toContain(
      '<a:tcPr xmlns:x="urn:test" x:marL="777" vert="wordArtVert" keep="CLEAR-UNDEFINED"/>',
    );
    expect(updated).toContain('<a:tc hMerge="1">');
    expect(updated).toContain(
      '<a:tcPr marB="76200" anchor="ctr" vert="wordArtVert" keep="MERGED"/>',
    );
    expect(updated).toContain('<a:bodyPr lIns="99999" tIns="88888" anchor="b" custom="BODY">');
    expect(updated).toContain(adjacentCell);

    table.setCellText(0, 2, 'Edited margins');
    table.setCellTextDirection(0, 1, 'wordArtVert');
    table.setCellTextFit(0, 0, 'resize');
    table.setCellVerticalAlignment(0, 1, 'bottom');
    table.setTransform({ x: inches(1) });
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain('<a:t>Edited margins</a:t>');
    expect(updated).toContain(
      '<a:tcPr anchor="b" marL="50800" marR="50800" marT="50800" marB="50800" vert="wordArtVert"/>',
    );
    expect(updated).toContain(
      '<a:tcPr marB=\'38100\' marT="12700" marR="25400" marL="50800" horzOverflow="clip">',
    );
    expect(updated).toContain('<a:bodyPr lIns="99999" tIns="88888" anchor="b" custom="BODY"><a:spAutoFit/>');
    expect(updated).toContain('<a:off x="914400" y="0"/>');
    expect(updated).toContain(adjacentCell);

    expect(table.rows[0]!.cells[1]!.margins).toEqual({
      top: 4,
      right: 4,
      bottom: 4,
      left: 4,
    });
    expect(table.rows[0]!.cells[3]!.margins).toEqual({ top: 4, left: 8 });
    expect(table.rows[0]!.cells[4]!.margins).toBeUndefined();
    expect(table.rows[0]!.cells[5]!.margins).toBeUndefined();
    expect(table.rows[0]!.cells[6]!.margins).toEqual({
      top: 3,
      right: 2,
      bottom: 4,
      left: 1,
    });
    expect(table.rows[0]!.cells[7]!.margins).toEqual({ bottom: 6 });

    const nullMargins = Object.assign(Object.create(null), {
      top: 1.500004,
      left: -2,
    });
    table.setCellMargins(0, 3, nullMargins);
    expect(table.rows[0]!.cells[3]!.margins).toEqual({ top: 1.5, left: -2 });
    nullMargins.top = 99;
    nullMargins.left = 99;
    expect(table.rows[0]!.cells[3]!.margins).toEqual({ top: 1.5, left: -2 });
    table.setCellMargins(0, 3, { top: 4, left: 8 });

    let marginAccessorCalls = 0;
    const accessorTuple = [1, 2, 3, 4];
    Object.defineProperty(accessorTuple, '0', {
      get() {
        marginAccessorCalls += 1;
        return 1;
      },
      enumerable: true,
      configurable: true,
    });
    const accessorNamed = { right: 2 };
    Object.defineProperty(accessorNamed, 'top', {
      get() {
        marginAccessorCalls += 1;
        return 1;
      },
      enumerable: true,
      configurable: true,
    });
    class MarginClass {
      top = 1;
    }
    const inheritedMargin = Object.assign(Object.create({ right: 2 }), { top: 1 });
    const symbolMargin = { top: 1, [Symbol('margin')]: 2 };
    const arraySubclass = new (class extends Array<number> {})(1, 2, 3, 4);

    const beforeInvalid = pkg.requirePart(part.uri).bytes;
    const invalidJournal = [...pkg.mutations];
    for (const value of [
      accessorTuple,
      accessorNamed,
      new MarginClass(),
      inheritedMargin,
      symbolMargin,
      arraySubclass,
    ]) {
      expect(() => table.setCellMargins(0, 0, value as never)).toThrow();
    }
    expect(marginAccessorCalls).toBe(0);
    for (const [row, column] of [[-1, 0], [0, -1], [0, 11], [1, 0]]) {
      expect(() => table.setCellMargins(row!, column!, 4)).toThrow(RangeError);
    }
    for (const column of [8, 9, 10]) {
      expect(() => table.setCellMargins(0, column, 4)).toThrow(ModelParseError);
    }
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);

    const beforeRollback = pkg.requirePart(part.uri).bytes;
    const rollbackJournal = [...pkg.mutations];
    const rollbackMargins = table.rows[0]!.cells.map(({ margins }) => margins);
    expect(() =>
      pkg.transaction(() => {
        table.setCellMargins(0, 0, 1);
        table.setCellMargins(0, 3, undefined);
        throw new Error('restore table cell margins');
      }),
    ).toThrow('restore table cell margins');
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(model.slides[1]!.shapes[2]).toBe(table);
    expect(table.rows[0]!.cells.map(({ margins }) => margins)).toEqual(rollbackMargins);
  });

  it('reads strict direct table-cell fills into detached snapshots', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const part = pkg.requirePart('/ppt/slides/slide1.xml');
    const cell = (properties: string, index: number): string =>
      `<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="111111"/></a:solidFill></a:rPr><a:t>Fill ${index}</a:t></a:r></a:p></a:txBody>${properties}</a:tc>`;
    const cells = [
      cell('<a:tcPr><a:noFill/></a:tcPr>', 0),
      cell('<a:tcPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:tcPr>', 1),
      cell('<a:tcPr><a:solidFill><a:schemeClr val="accent1"><a:alpha val="75000"/></a:schemeClr></a:solidFill></a:tcPr>', 2),
      cell('<a:tcPr><a:solidFill><a:srgbClr val="00FF00"><a:alpha val="100000"/></a:srgbClr></a:solidFill></a:tcPr>', 3),
      cell('<a:tcPr><a:solidFill><a:srgbClr val="0000FF"><a:alpha val="0"/></a:srgbClr></a:solidFill></a:tcPr>', 4),
      cell('<a:tcPr><a:solidFill><a:srgbClr val="abcdef"><a:alpha val="075000"/></a:srgbClr></a:solidFill></a:tcPr>', 5),
      cell('<a:tcPr/>', 6),
      cell('<a:tcPr><a:lnL><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:lnL></a:tcPr>', 7),
      cell('<a:tcPr><a:gradFill><a:gsLst/></a:gradFill></a:tcPr>', 8),
      cell('<a:tcPr><a:pattFill prst="pct5"/></a:tcPr>', 9),
      cell('<a:tcPr><a:noFill/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:tcPr>', 10),
      cell('<a:tcPr><a:solidFill><a:srgbClr val="FFF"/></a:solidFill></a:tcPr>', 11),
      cell('<a:tcPr><a:solidFill><a:schemeClr val="unknown"/></a:solidFill></a:tcPr>', 12),
      cell('<a:tcPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val=""/></a:srgbClr></a:solidFill></a:tcPr>', 13),
      cell('<a:tcPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="1.5"/></a:srgbClr></a:solidFill></a:tcPr>', 14),
      cell('<a:tcPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="-1"/></a:srgbClr></a:solidFill></a:tcPr>', 15),
      cell('<a:tcPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="100001"/></a:srgbClr></a:solidFill></a:tcPr>', 16),
      cell('<a:tcPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/><a:alpha val="75000"/></a:srgbClr></a:solidFill></a:tcPr>', 17),
      cell('<a:tcPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="75000"/><a:tint val="50000"/></a:srgbClr></a:solidFill></a:tcPr>', 18),
      cell('<a:tcPr xmlns:x="urn:test"><a:solidFill><a:srgbClr x:val="FF0000"/></a:solidFill></a:tcPr>', 19),
      cell('<a:tcPr xmlns:x="urn:test"><x:solidFill><x:srgbClr val="FF0000"/></x:solidFill></a:tcPr>', 20),
      cell('<a:tcPr><a:noFill custom="INVALID"/></a:tcPr>', 21),
      cell('<a:tcPr><a:solidFill custom="INVALID"><a:srgbClr val="FF0000"/></a:solidFill></a:tcPr>', 22),
      cell('<a:tcPr xmlns:x="urn:test"><a:solidFill><a:srgbClr val="FF0000"><a:alpha x:val="75000"/></a:srgbClr></a:solidFill></a:tcPr>', 23),
      cell('<a:tcPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="75000"><a:ext/></a:alpha></a:srgbClr></a:solidFill></a:tcPr>', 24),
      cell('<a:tcPr><a:noFill/></a:tcPr><a:tcPr keep="SECOND"/>', 25),
      cell('', 26),
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(/<a:tr>.*?<\/a:tr>/, `<a:tr>${cells}</a:tr>`),
      part.contentType,
    );
    const table = new PresentationModel(pkg).slides[1]!.shapes[2] as TableModel;
    const journal = [...pkg.mutations];

    const snapshot = table.rows;
    expect(snapshot[0]!.cells.map(({ fill }) => fill)).toEqual([
      { kind: 'none' },
      { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } },
      {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 25,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: '00FF00' },
        transparency: 0,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: '0000FF' },
        transparency: 100,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'ABCDEF' },
        transparency: 25,
      },
      ...Array(21).fill(undefined),
    ]);
    expect(pkg.mutations).toEqual(journal);

    const mutable = snapshot[0]!.cells[2]!.fill as {
      kind: string;
      color: { kind: string; value: string };
      transparency?: number;
    };
    mutable.kind = 'none';
    mutable.color.value = 'FFFFFF';
    mutable.transparency = 100;
    expect(table.rows[0]!.cells[2]!.fill).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 25,
    });
    expect(pkg.mutations).toEqual(journal);
  });

  it('losslessly replaces one table-cell fill and rolls back atomically', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const part = pkg.requirePart('/ppt/slides/slide1.xml');
    const textBody = (bodyProperties: string, text: string): string =>
      `<a:txBody>${bodyProperties}<a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="222222"/></a:solidFill></a:rPr><a:t>${text}</a:t></a:r></a:p></a:txBody>`;
    const cell = (body: string, properties: string, attributes = ''): string =>
      `<a:tc${attributes}>${body}${properties}</a:tc>`;
    const adjacentCell = cell(
      textBody('<a:bodyPr lIns="12700" keep="ADJACENT"><a:spAutoFit/></a:bodyPr>', 'Adjacent'),
      '<a:tcPr marL="12700" marR="25400" marT="38100" marB="50800" anchor="b" vert="horz" keep="ADJACENT-TCPR"><a:lnL w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:lnL><a:solidFill><a:srgbClr val="70AD47"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:tcPr>',
    );
    const cells = [
      cell(
        textBody('<a:bodyPr custom="SAME"><a:normAutofit fontScale="85000"/></a:bodyPr>', 'Same fill'),
        '<a:tcPr marL="12700" anchor="ctr" vert="vert270" custom="KEEP"><a:solidFill><a:srgbClr val="abcdef"><a:alpha val="075000"/></a:srgbClr></a:solidFill><x:keep xmlns:x="urn:test">TCPR</x:keep></a:tcPr>',
      ),
      cell(
        textBody('<a:bodyPr/>', 'Alternate prefix'),
        '<q:tcPr xmlns:q="a" anchor="t"/>',
      ),
      cell(
        textBody('<a:bodyPr/>', 'Replace no fill'),
        '<a:tcPr anchor="b"><a:lnL w="12700"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:lnL><a:noFill/><a:extLst><a:ext uri="KEEP"/></a:extLst></a:tcPr>',
      ),
      cell(
        textBody('<a:bodyPr/>', 'Replace gradient'),
        '<a:tcPr keep="GRADIENT"><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst></a:gradFill></a:tcPr>',
      ),
      cell(
        textBody('<a:bodyPr/>', 'Clear solid'),
        '<a:tcPr keep="CLEAR"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><x:keep xmlns:x="urn:test">CLEAR</x:keep></a:tcPr>',
      ),
      cell(
        textBody('<a:bodyPr/>', 'Canonicalize malformed'),
        '<a:tcPr keep="MALFORMED"><a:solidFill><a:srgbClr val="FFF"><a:alpha val="invalid"/></a:srgbClr></a:solidFill></a:tcPr>',
      ),
      adjacentCell,
      cell(
        textBody('<a:bodyPr/>', 'Merged'),
        '<a:tcPr/>',
        ' hMerge="1"',
      ),
      cell(
        textBody('<a:bodyPr/>', 'Multiple fills'),
        '<a:tcPr><a:noFill/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:tcPr>',
      ),
      cell(textBody('<a:bodyPr/>', 'Repeated tcPr'), '<a:tcPr/><a:tcPr keep="SECOND"/>'),
      cell(textBody('<a:bodyPr/>', 'Missing tcPr'), ''),
      cell(
        textBody('<a:bodyPr/>', 'Wrong prefix'),
        '<a:tcPr xmlns:x="urn:test"><x:solidFill><x:srgbClr val="999999"/></x:solidFill><a:extLst><a:ext uri="LAST"/></a:extLst></a:tcPr>',
      ),
    ].join('');
    const source = new TextDecoder().decode(part.bytes)
      .replace(
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr><a:graphic>',
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="2000"/></p:xfrm><a:graphic>',
      )
      .replace(/<a:tr>.*?<\/a:tr>/, `<a:tr>${cells}</a:tr>`);
    pkg.setPart(part.uri, source, part.contentType);
    const model = new PresentationModel(pkg);
    const table = model.slides[1]!.shapes[2] as TableModel;

    const noOpJournal = [...pkg.mutations];
    table.setCellFill(0, 0, {
      kind: 'solid',
      color: { kind: 'srgb', value: '#ABCDEF' },
      transparency: 25,
    });
    expect(pkg.mutations).toEqual(noOpJournal);
    expect(new TextDecoder().decode(pkg.requirePart(part.uri).bytes)).toContain(
      '<a:srgbClr val="abcdef"><a:alpha val="075000"/></a:srgbClr>',
    );

    table.setCellFill(0, 1, {
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF0000' },
    });
    table.setCellFill(0, 2, {
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 33.333,
    });
    table.setCellFill(0, 3, { kind: 'none' });
    table.setCellFill(0, 4, undefined);
    table.setCellFill(0, 5, {
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent3' },
      transparency: 0,
    });
    table.setCellFill(0, 7, {
      kind: 'solid',
      color: { kind: 'srgb', value: '0000FF' },
      transparency: 100,
    });
    const nullPrototypeColor = Object.assign(Object.create(null), {
      kind: 'srgb',
      value: '#445566',
    });
    const nullPrototypeFill = Object.assign(Object.create(null), {
      kind: 'solid',
      color: nullPrototypeColor,
    });
    table.setCellFill(0, 11, nullPrototypeFill);
    let updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain(
      '<q:tcPr xmlns:q="a" anchor="t"><q:solidFill><q:srgbClr val="FF0000"/></q:solidFill></q:tcPr>',
    );
    expect(updated).toContain(
      '<a:lnL w="12700"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:lnL><a:solidFill><a:schemeClr val="accent2"><a:alpha val="66667"/></a:schemeClr></a:solidFill><a:extLst>',
    );
    expect(updated).toContain('<a:tcPr keep="GRADIENT"><a:noFill/></a:tcPr>');
    expect(updated).toContain(
      '<a:tcPr keep="CLEAR"><x:keep xmlns:x="urn:test">CLEAR</x:keep></a:tcPr>',
    );
    expect(updated).toContain(
      '<a:tcPr keep="MALFORMED"><a:solidFill><a:schemeClr val="accent3"><a:alpha val="100000"/></a:schemeClr></a:solidFill></a:tcPr>',
    );
    expect(updated).toContain('<a:tc hMerge="1">');
    expect(updated).toContain(
      '<a:tcPr><a:solidFill><a:srgbClr val="0000FF"><a:alpha val="0"/></a:srgbClr></a:solidFill></a:tcPr>',
    );
    expect(updated).toContain(
      '<x:solidFill><x:srgbClr val="999999"/></x:solidFill><a:solidFill><a:srgbClr val="445566"/></a:solidFill><a:extLst>',
    );
    expect(updated).toContain(adjacentCell);

    table.setCellTextFit(0, 0, 'resize');
    table.setCellTextDirection(0, 1, 'wordArtVert');
    table.setCellVerticalAlignment(0, 1, 'bottom');
    table.setCellMargins(0, 2, { top: 4, left: 8 });
    table.setCellText(0, 2, 'Edited fill');
    table.setTransform({ x: inches(1) });
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain('<a:t>Edited fill</a:t>');
    expect(updated).toContain('<a:bodyPr custom="SAME"><a:spAutoFit/></a:bodyPr>');
    expect(updated).toContain('<a:srgbClr val="abcdef"><a:alpha val="075000"/></a:srgbClr>');
    expect(updated).toContain(
      '<q:tcPr xmlns:q="a" anchor="b" vert="wordArtVert"><q:solidFill><q:srgbClr val="FF0000"/></q:solidFill></q:tcPr>',
    );
    expect(updated).toContain(
      '<a:tcPr anchor="b" marL="101600" marT="50800"><a:lnL w="12700">',
    );
    expect(updated).toContain('<a:off x="914400" y="0"/>');
    expect(updated).toContain(adjacentCell);

    expect(table.rows[0]!.cells.map(({ fill }) => fill)).toEqual([
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'ABCDEF' },
        transparency: 25,
      },
      { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } },
      {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 33.333,
      },
      { kind: 'none' },
      undefined,
      {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent3' },
        transparency: 0,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: '70AD47' },
        transparency: 50,
      },
      {
        kind: 'solid',
        color: { kind: 'srgb', value: '0000FF' },
        transparency: 100,
      },
      undefined,
      undefined,
      undefined,
      { kind: 'solid', color: { kind: 'srgb', value: '445566' } },
    ]);

    const beforeInvalid = pkg.requirePart(part.uri).bytes;
    const invalidJournal = [...pkg.mutations];
    for (const [row, column] of [[-1, 0], [0, -1], [0, 12], [1, 0]]) {
      expect(() => table.setCellFill(row!, column!, { kind: 'none' })).toThrow(RangeError);
    }
    for (const column of [8, 9, 10]) {
      expect(() => table.setCellFill(0, column, { kind: 'none' })).toThrow(ModelParseError);
    }
    const accessorFill = {};
    const accessorColor = { kind: 'srgb' };
    let fillAccessorCalls = 0;
    Object.defineProperty(accessorFill, 'kind', {
      get() {
        fillAccessorCalls += 1;
        return 'none';
      },
      enumerable: true,
      configurable: true,
    });
    Object.defineProperty(accessorColor, 'value', {
      get() {
        fillAccessorCalls += 1;
        return 'FF0000';
      },
      enumerable: true,
      configurable: true,
    });
    class ExoticFill {
      kind = 'none';
    }
    const invalidValues = [
      null,
      false,
      '',
      [],
      {},
      { kind: 'none', color: { kind: 'srgb', value: 'FF0000' } },
      { kind: 'unknown' },
      { kind: 'solid' },
      { kind: 'solid', color: { kind: 'srgb', value: 'FFF' } },
      { kind: 'solid', color: { kind: 'scheme', value: 'unknown' } },
      { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: -1 },
      { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: 101 },
      { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' }, transparency: Number.NaN },
      accessorFill,
      Object.create({ kind: 'none' }),
      new ExoticFill(),
      Object.assign({ kind: 'none' }, { [Symbol('extra')]: true }),
      { kind: 'solid', color: accessorColor },
      Symbol('table cell fill'),
    ];
    for (const value of invalidValues) {
      expect(() => table.setCellFill(0, 0, value as never)).toThrow();
    }
    expect(fillAccessorCalls).toBe(0);
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);

    const beforeRollback = pkg.requirePart(part.uri).bytes;
    const rollbackJournal = [...pkg.mutations];
    const rollbackFills = table.rows[0]!.cells.map(({ fill }) => fill);
    expect(() =>
      pkg.transaction(() => {
        table.setCellFill(0, 0, { kind: 'none' });
        table.setCellFill(0, 4, {
          kind: 'solid',
          color: { kind: 'srgb', value: 'FFFFFF' },
        });
        throw new Error('restore table cell fills');
      }),
    ).toThrow('restore table cell fills');
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(model.slides[1]!.shapes[2]).toBe(table);
    expect(table.rows[0]!.cells.map(({ fill }) => fill)).toEqual(rollbackFills);
  });

  it('reads strict direct table-cell borders into detached partial snapshots', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const part = pkg.requirePart('/ppt/slides/slide1.xml');
    const cell = (properties: string, index: number): string =>
      `<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Border ${index}</a:t></a:r></a:p></a:txBody>${properties}</a:tc>`;
    const cells = [
      cell(
        '<a:tcPr><a:lnL w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnL><a:lnR w="12700"><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></a:lnR><a:lnT w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:lnT><a:lnB w="25400"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:prstDash val="sysDash"/></a:lnB></a:tcPr>',
        0,
      ),
      cell(
        '<a:tcPr><a:lnL w="00012700"><a:solidFill><a:srgbClr val="abcdef"/></a:solidFill></a:lnL><a:lnR w="20116800"><a:solidFill><a:schemeClr val="accent6"/></a:solidFill></a:lnR><a:lnT w="0"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:lnT></a:tcPr>',
        1,
      ),
      cell(
        '<a:tcPr><a:lnL w="-1"><a:noFill/></a:lnL><a:lnR w="1.5"><a:noFill/></a:lnR><a:lnT w="20116801"><a:noFill/></a:lnT><a:lnB w=""><a:noFill/></a:lnB></a:tcPr>',
        2,
      ),
      cell(
        '<a:tcPr><a:lnL w="12700"><a:noFill/></a:lnL><a:lnL w="25400"><a:noFill/></a:lnL><a:lnT w="12700"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:lnT></a:tcPr>',
        3,
      ),
      cell(
        '<a:tcPr><a:lnL w="12700"><a:gradFill><a:gsLst/></a:gradFill></a:lnL><a:lnR w="12700"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:prstDash val="sysDash"/></a:lnR></a:tcPr>',
        4,
      ),
      cell(
        '<a:tcPr><a:lnT w="12700"><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:lnT><a:lnB w="12700"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:prstDash val="dash"/></a:lnB></a:tcPr>',
        5,
      ),
      cell(
        '<a:tcPr><a:lnT w="0"><a:noFill custom="INVALID"/></a:lnT><a:lnB w="0"><a:noFill/><a:round/></a:lnB></a:tcPr>',
        6,
      ),
      cell(
        '<a:tcPr><a:lnL w="12700"><a:solidFill><a:srgbClr val="778899"/></a:solidFill></a:lnL><a:lnT w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:headEnd type="arrow" w="med" len="med"/></a:lnT></a:tcPr>',
        7,
      ),
      cell(
        '<a:tcPr xmlns:x="urn:test"><x:lnL w="12700"><x:noFill/></x:lnL><a:lnR w="12700"><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill></a:lnR></a:tcPr>',
        8,
      ),
      cell(
        '<a:tcPr><a:lnTlToBr w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:lnTlToBr><a:lnBlToTr w="12700"><a:noFill/></a:lnBlToTr><a:lnB w="12700"><a:noFill/></a:lnB></a:tcPr>',
        9,
      ),
      cell('<a:tcPr><a:lnL w="12700"><a:noFill/></a:lnL></a:tcPr><a:tcPr/>', 10),
      cell('', 11),
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(/<a:tr>.*?<\/a:tr>/, `<a:tr>${cells}</a:tr>`),
      part.contentType,
    );
    const table = new PresentationModel(pkg).slides[1]!.shapes[2] as TableModel;
    const journal = [...pkg.mutations];

    const snapshot = table.rows;
    expect(snapshot[0]!.cells.map(({ borders }) => borders)).toEqual([
      {
        top: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent1' },
          width: 1.5,
          style: 'solid',
        },
        right: {
          kind: 'line',
          color: { kind: 'srgb', value: 'FF0000' },
          width: 1,
        },
        bottom: {
          kind: 'line',
          color: { kind: 'srgb', value: '0000FF' },
          width: 2,
          style: 'dash',
        },
        left: { kind: 'none' },
      },
      {
        top: {
          kind: 'line',
          color: { kind: 'srgb', value: '112233' },
          width: 0,
        },
        right: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent6' },
          width: 1584,
        },
        left: {
          kind: 'line',
          color: { kind: 'srgb', value: 'ABCDEF' },
          width: 1,
        },
      },
      undefined,
      {
        top: {
          kind: 'line',
          color: { kind: 'srgb', value: '445566' },
          width: 1,
        },
      },
      {
        right: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent2' },
          width: 1,
          style: 'dash',
        },
      },
      undefined,
      undefined,
      {
        left: {
          kind: 'line',
          color: { kind: 'srgb', value: '778899' },
          width: 1,
        },
      },
      {
        right: {
          kind: 'line',
          color: { kind: 'srgb', value: 'AABBCC' },
          width: 1,
        },
      },
      { bottom: { kind: 'none' } },
      undefined,
      undefined,
    ]);
    expect(pkg.mutations).toEqual(journal);

    const mutable = snapshot[0]!.cells[0]!.borders as {
      top?: { kind: string; color: { kind: string; value: string }; width: number };
      left?: { kind: string };
    };
    mutable.top!.kind = 'none';
    mutable.top!.color.value = 'FFFFFF';
    mutable.top!.width = 99;
    mutable.left!.kind = 'line';
    expect(table.rows[0]!.cells[0]!.borders).toEqual({
      top: {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent1' },
        width: 1.5,
        style: 'solid',
      },
      right: {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        width: 1,
      },
      bottom: {
        kind: 'line',
        color: { kind: 'srgb', value: '0000FF' },
        width: 2,
        style: 'dash',
      },
      left: { kind: 'none' },
    });
    expect(pkg.mutations).toEqual(journal);
  });

  it('losslessly replaces table-cell borders and rolls back atomically', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const part = pkg.requirePart('/ppt/slides/slide1.xml');
    const textBody = (bodyProperties: string, text: string): string =>
      `<a:txBody>${bodyProperties}<a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="222222"/></a:solidFill></a:rPr><a:t>${text}</a:t></a:r></a:p></a:txBody>`;
    const cell = (body: string, properties: string, attributes = ''): string =>
      `<a:tc${attributes}>${body}${properties}</a:tc>`;
    const solidLine = (
      tag: 'lnL' | 'lnR' | 'lnT' | 'lnB',
      color = 'abcdef',
      width = '012700',
      dash = 'solid',
    ): string =>
      `<a:${tag} w="${width}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:prstDash val="${dash}"/><a:round/><a:headEnd type="none" w="med" len="med"/><a:tailEnd type="none" w="med" len="med"/></a:${tag}>`;
    const sameBorders = [
      solidLine('lnL'),
      solidLine('lnR'),
      solidLine('lnT'),
      solidLine('lnB'),
    ].join('');
    const adjacentCell = cell(
      textBody('<a:bodyPr lIns="12700" keep="ADJACENT"><a:spAutoFit/></a:bodyPr>', 'Adjacent'),
      `<a:tcPr marL="12700" marR="25400" marT="38100" marB="50800" anchor="b" vert="horz" keep="ADJACENT-TCPR">${solidLine('lnL', '333333', '25400')}<a:solidFill><a:srgbClr val="70AD47"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:tcPr>`,
    );
    const cells = [
      cell(
        textBody('<a:bodyPr custom="SAME"><a:normAutofit fontScale="85000"/></a:bodyPr>', 'Same borders'),
        `<a:tcPr marL="12700" anchor="ctr" vert="vert270" custom="KEEP">${sameBorders}<a:solidFill><a:schemeClr val="accent6"/></a:solidFill><x:keep xmlns:x="urn:test">TCPR</x:keep></a:tcPr>`,
      ),
      cell(textBody('<a:bodyPr/>', 'Alternate prefix'), '<q:tcPr xmlns:q="a" anchor="t"/>'),
      cell(
        textBody('<a:bodyPr/>', 'Partial replacement'),
        '<a:tcPr anchor="b"><a:lnL w="12700"><a:gradFill><a:gsLst/></a:gradFill></a:lnL><a:lnR w="0"><a:noFill/></a:lnR><a:lnTlToBr w="12700"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:lnTlToBr><a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill><a:cell3D prstMaterial="flat"/><a:extLst><a:ext uri="KEEP"/></a:extLst></a:tcPr>',
      ),
      cell(
        textBody('<a:bodyPr/>', 'Clear object'),
        `<a:tcPr keep="CLEAR-OBJECT">${sameBorders}<x:keep xmlns:x="urn:test">OBJECT</x:keep></a:tcPr>`,
      ),
      cell(
        textBody('<a:bodyPr/>', 'Clear undefined'),
        `<a:tcPr keep="CLEAR-UNDEFINED">${sameBorders}<x:keep xmlns:x="urn:test">UNDEFINED</x:keep></a:tcPr>`,
      ),
      cell(
        textBody('<a:bodyPr/>', 'Malformed replacement'),
        '<a:tcPr keep="MALFORMED"><a:lnT w="invalid"><a:solidFill><a:srgbClr val="FFF"/></a:solidFill><a:prstDash val="dot"/></a:lnT></a:tcPr>',
      ),
      adjacentCell,
      cell(textBody('<a:bodyPr/>', 'Merged'), '<a:tcPr/>', ' hMerge="1"'),
      cell(
        textBody('<a:bodyPr/>', 'Repeated side'),
        '<a:tcPr><a:lnL w="0"><a:noFill/></a:lnL><a:lnL w="0"><a:noFill/></a:lnL></a:tcPr>',
      ),
      cell(textBody('<a:bodyPr/>', 'Repeated tcPr'), '<a:tcPr/><a:tcPr keep="SECOND"/>'),
      cell(textBody('<a:bodyPr/>', 'Missing tcPr'), ''),
      cell(
        textBody('<a:bodyPr/>', 'Wrong prefix'),
        '<a:tcPr xmlns:x="urn:test"><x:lnL w="0"><x:noFill/></x:lnL><a:extLst><a:ext uri="LAST"/></a:extLst></a:tcPr>',
      ),
    ].join('');
    const source = new TextDecoder().decode(part.bytes)
      .replace(
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr><a:graphic>',
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="2000"/></p:xfrm><a:graphic>',
      )
      .replace(/<a:tr>.*?<\/a:tr>/, `<a:tr>${cells}</a:tr>`);
    pkg.setPart(part.uri, source, part.contentType);
    const model = new PresentationModel(pkg);
    const table = model.slides[1]!.shapes[2] as TableModel;

    const noOpBytes = pkg.requirePart(part.uri).bytes;
    const noOpJournal = [...pkg.mutations];
    table.setCellBorders(0, 0, {
      kind: 'line',
      color: { kind: 'srgb', value: '#ABCDEF' },
      width: 1,
      style: 'solid',
    });
    expect(pkg.requirePart(part.uri).bytes).toEqual(noOpBytes);
    expect(pkg.mutations).toEqual(noOpJournal);
    expect(new TextDecoder().decode(pkg.requirePart(part.uri).bytes)).toContain(sameBorders);

    table.setCellBorders(0, 1, [
      {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent2' },
        width: 1.5,
        style: 'dash',
      },
      {
        kind: 'line',
        color: { kind: 'srgb', value: '00FF00' },
        width: 0,
      },
      { kind: 'none' },
      undefined,
    ]);
    table.setCellBorders(0, 2, {
      top: {
        kind: 'line',
        color: { kind: 'srgb', value: '112233' },
        width: 2,
      },
      left: { kind: 'none' },
    });
    table.setCellBorders(0, 3, {});
    table.setCellBorders(0, 4, undefined);
    table.setCellBorders(0, 5, {
      kind: 'line',
      color: { kind: 'scheme', value: 'accent3' },
      width: 2.25,
      style: 'dash',
    });
    table.setCellBorders(0, 7, { kind: 'none' });
    table.setCellBorders(0, 11, {
      kind: 'line',
      color: { kind: 'srgb', value: '445566' },
      width: 0.5,
    });

    let updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain(
      '<q:tcPr xmlns:q="a" anchor="t"><q:lnR w="0" cap="flat" cmpd="sng" algn="ctr"><q:solidFill><q:srgbClr val="00FF00"/></q:solidFill><q:round/><q:headEnd type="none" w="med" len="med"/><q:tailEnd type="none" w="med" len="med"/></q:lnR><q:lnT w="19050" cap="flat" cmpd="sng" algn="ctr"><q:solidFill><q:schemeClr val="accent2"/></q:solidFill><q:prstDash val="sysDash"/>',
    );
    expect(updated).toContain(
      '</q:lnT><q:lnB w="0" cap="flat" cmpd="sng" algn="ctr"><q:noFill/></q:lnB></q:tcPr>',
    );
    expect(updated).toContain(
      '<a:tcPr anchor="b"><a:lnL w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/></a:lnL><a:lnT w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:round/>',
    );
    expect(updated).toContain(
      '</a:lnT><a:lnTlToBr w="12700"><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:lnTlToBr><a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill><a:cell3D prstMaterial="flat"/><a:extLst>',
    );
    expect(updated).toContain(
      '<a:tcPr keep="CLEAR-OBJECT"><x:keep xmlns:x="urn:test">OBJECT</x:keep></a:tcPr>',
    );
    expect(updated).toContain(
      '<a:tcPr keep="CLEAR-UNDEFINED"><x:keep xmlns:x="urn:test">UNDEFINED</x:keep></a:tcPr>',
    );
    expect(updated).toContain('<a:tc hMerge="1">');
    expect(updated).toContain(
      '<x:lnL w="0"><x:noFill/></x:lnL><a:lnL w="6350" cap="flat" cmpd="sng" algn="ctr">',
    );
    expect(updated).toContain(
      '</a:lnB><a:extLst><a:ext uri="LAST"/></a:extLst>',
    );
    expect(updated).toContain(adjacentCell);

    table.setCellTextFit(0, 0, 'resize');
    table.setCellTextDirection(0, 1, 'wordArtVert');
    table.setCellVerticalAlignment(0, 1, 'bottom');
    table.setCellMargins(0, 2, { top: 4, left: 8 });
    table.setCellFill(0, 2, {
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent5' },
    });
    table.setCellText(0, 2, 'Edited borders');
    table.setTransform({ x: inches(1) });
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain('<a:t>Edited borders</a:t>');
    expect(updated).toContain('<a:off x="914400" y="0"/>');
    expect(updated).toContain('<a:lnTlToBr w="12700">');
    expect(updated).toContain('<a:solidFill><a:schemeClr val="accent5"/></a:solidFill>');
    expect(updated).toContain(adjacentCell);

    const line = (
      color: { kind: 'srgb' | 'scheme'; value: string },
      width: number,
      style?: 'solid' | 'dash',
    ) => ({ kind: 'line' as const, color, width, ...(style ? { style } : {}) });
    expect(table.rows[0]!.cells.map(({ borders }) => borders)).toEqual([
      {
        top: line({ kind: 'srgb', value: 'ABCDEF' }, 1, 'solid'),
        right: line({ kind: 'srgb', value: 'ABCDEF' }, 1, 'solid'),
        bottom: line({ kind: 'srgb', value: 'ABCDEF' }, 1, 'solid'),
        left: line({ kind: 'srgb', value: 'ABCDEF' }, 1, 'solid'),
      },
      {
        top: line({ kind: 'scheme', value: 'accent2' }, 1.5, 'dash'),
        right: line({ kind: 'srgb', value: '00FF00' }, 0),
        bottom: { kind: 'none' },
      },
      {
        top: line({ kind: 'srgb', value: '112233' }, 2),
        left: { kind: 'none' },
      },
      undefined,
      undefined,
      {
        top: line({ kind: 'scheme', value: 'accent3' }, 2.25, 'dash'),
        right: line({ kind: 'scheme', value: 'accent3' }, 2.25, 'dash'),
        bottom: line({ kind: 'scheme', value: 'accent3' }, 2.25, 'dash'),
        left: line({ kind: 'scheme', value: 'accent3' }, 2.25, 'dash'),
      },
      { left: line({ kind: 'srgb', value: '333333' }, 2, 'solid') },
      {
        top: { kind: 'none' },
        right: { kind: 'none' },
        bottom: { kind: 'none' },
        left: { kind: 'none' },
      },
      undefined,
      undefined,
      undefined,
      {
        top: line({ kind: 'srgb', value: '445566' }, 0.5),
        right: line({ kind: 'srgb', value: '445566' }, 0.5),
        bottom: line({ kind: 'srgb', value: '445566' }, 0.5),
        left: line({ kind: 'srgb', value: '445566' }, 0.5),
      },
    ]);

    const nullPrototypeColor = Object.assign(Object.create(null) as Record<string, unknown>, {
      kind: 'srgb',
      value: '#112233',
    });
    const nullPrototypeLine = Object.assign(Object.create(null) as Record<string, unknown>, {
      kind: 'line',
      color: nullPrototypeColor,
      width: 2,
    });
    const nullPrototypeNone = Object.assign(Object.create(null) as Record<string, unknown>, {
      kind: 'none',
    });
    const nullPrototypeNamed = Object.assign(Object.create(null) as Record<string, unknown>, {
      top: nullPrototypeLine,
      left: nullPrototypeNone,
    });
    const beforeNullPrototype = pkg.requirePart(part.uri).bytes;
    const nullPrototypeJournal = [...pkg.mutations];
    table.setCellBorders(0, 2, nullPrototypeNamed as never);
    expect(table.rows[0]!.cells[2]!.borders).toEqual({
      top: line({ kind: 'srgb', value: '112233' }, 2),
      left: { kind: 'none' },
    });
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeNullPrototype);
    expect(pkg.mutations).toEqual(nullPrototypeJournal);

    const beforeInvalid = pkg.requirePart(part.uri).bytes;
    const invalidJournal = [...pkg.mutations];
    for (const [row, column] of [[-1, 0], [0, -1], [0, 12], [1, 0]]) {
      expect(() => table.setCellBorders(row!, column!, { kind: 'none' })).toThrow(RangeError);
    }
    for (const column of [8, 9, 10]) {
      expect(() => table.setCellBorders(0, column, { kind: 'none' })).toThrow(ModelParseError);
    }
    const sparse = Array(4);
    sparse[0] = { kind: 'none' };
    const extraTuple = Object.assign(
      [{ kind: 'none' }, undefined, undefined, undefined],
      { extra: true },
    );
    let borderAccessorCalls = 0;
    const accessorBorder: Record<string, unknown> = {};
    Object.defineProperty(accessorBorder, 'kind', {
      get() {
        borderAccessorCalls += 1;
        return 'none';
      },
      enumerable: true,
    });
    const accessorNamed: Record<string, unknown> = {};
    Object.defineProperty(accessorNamed, 'top', {
      get() {
        borderAccessorCalls += 1;
        return { kind: 'none' };
      },
      enumerable: true,
    });
    const accessorTuple: unknown[] = [undefined, undefined, undefined, undefined];
    Object.defineProperty(accessorTuple, '0', {
      get() {
        borderAccessorCalls += 1;
        return { kind: 'none' };
      },
      enumerable: true,
    });
    const accessorColor: Record<string, unknown> = { kind: 'srgb' };
    Object.defineProperty(accessorColor, 'value', {
      get() {
        borderAccessorCalls += 1;
        return 'FF0000';
      },
      enumerable: true,
    });
    class BorderValue {
      readonly kind = 'none';
    }
    class ColorValue {
      readonly kind = 'srgb';
      readonly value = 'FF0000';
    }
    class BorderTuple extends Array<unknown> {}
    const subclassTuple = new BorderTuple();
    subclassTuple.push({ kind: 'none' }, undefined, undefined, undefined);
    const symbol = Symbol('extra');
    const invalidValues = [
      null,
      false,
      '',
      [],
      [{ kind: 'none' }, undefined, undefined],
      sparse,
      extraTuple,
      accessorBorder,
      accessorNamed,
      accessorTuple,
      { kind: 'line', color: accessorColor, width: 1 },
      Object.create({ kind: 'none' }),
      Object.create({ top: { kind: 'none' } }),
      new BorderValue(),
      subclassTuple,
      { top: [{ kind: 'none' }] },
      Object.assign({ kind: 'none' }, { [symbol]: true }),
      { top: Object.assign({ kind: 'none' }, { [symbol]: true }) },
      {
        kind: 'line',
        color: Object.create({ kind: 'srgb', value: 'FF0000' }),
        width: 1,
      },
      { kind: 'line', color: new ColorValue(), width: 1 },
      {
        kind: 'line',
        color: Object.assign({ kind: 'srgb', value: 'FF0000' }, { [symbol]: true }),
        width: 1,
      },
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
      expect(() => table.setCellBorders(0, 0, value as never)).toThrow();
    }
    expect(borderAccessorCalls).toBe(0);
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);

    const beforeRollback = pkg.requirePart(part.uri).bytes;
    const rollbackJournal = [...pkg.mutations];
    const rollbackBorders = table.rows[0]!.cells.map(({ borders }) => borders);
    expect(() =>
      pkg.transaction(() => {
        table.setCellBorders(0, 0, { kind: 'none' });
        table.setCellBorders(0, 4, {
          kind: 'line',
          color: { kind: 'srgb', value: 'FFFFFF' },
          width: 1,
        });
        throw new Error('restore table cell borders');
      }),
    ).toThrow('restore table cell borders');
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(model.slides[1]!.shapes[2]).toBe(table);
    expect(table.rows[0]!.cells.map(({ borders }) => borders)).toEqual(rollbackBorders);
  });

  it('reads only exact direct table-cell vertical alignments into detached snapshots', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const part = pkg.requirePart('/ppt/slides/slide1.xml');
    const cell = (
      properties: string,
      index: number,
      bodyProperties = '<a:bodyPr/>',
    ): string =>
      `<a:tc><a:txBody>${bodyProperties}<a:p><a:r><a:t>Alignment ${index}</a:t></a:r></a:p></a:txBody>${properties}</a:tc>`;
    const cells = [
      cell('<a:tcPr anchor="t" vert="horz"/>', 0),
      cell('<a:tcPr anchor="ctr" vert="horz"/>', 1),
      cell('<a:tcPr anchor="b" vert="horz"/>', 2),
      cell('<a:tcPr vert="horz"/>', 3),
      cell('<a:tcPr anchor="" vert="horz"/>', 4),
      cell('<a:tcPr anchor="T" vert="horz"/>', 5),
      cell('<a:tcPr anchor=" ctr " vert="horz"/>', 6),
      cell('<a:tcPr anchor="top" vert="horz"/>', 7),
      cell('<a:tcPr anchor="middle" vert="horz"/>', 8),
      cell('<a:tcPr anchor="bottom" vert="horz"/>', 9),
      cell('<a:tcPr anchor="just" vert="horz"/>', 10),
      cell('<a:tcPr anchor="dist" vert="horz"/>', 11),
      cell('<a:tcPr anchor="unknown" vert="horz"/>', 12),
      cell('<a:tcPr xmlns:x="urn:test" x:anchor="t" vert="horz"/>', 13),
      cell('<a:tcPr anchor="t" anchor="b" vert="horz"/>', 14),
      cell('<a:tcPr vert="horz"/>', 15, '<a:bodyPr anchor="b"/>'),
      cell('<a:tcPr anchor="t"/><a:tcPr keep="SECOND"/>', 16),
      cell('', 17),
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(/<a:tr>.*?<\/a:tr>/, `<a:tr>${cells}</a:tr>`),
      part.contentType,
    );
    const table = new PresentationModel(pkg).slides[1]!.shapes[2] as TableModel;
    const journal = [...pkg.mutations];

    const snapshot = table.rows;
    const alignments = snapshot[0]!.cells.map(({ verticalAlignment }) => verticalAlignment);
    expect(alignments.slice(0, 3)).toEqual(['top', 'middle', 'bottom']);
    expect(alignments.slice(3)).toEqual(Array(15).fill(undefined));
    expect(snapshot[0]!.cells.slice(0, 16).every(({ textDirection }) => textDirection === 'horz')).toBe(true);
    expect(snapshot[0]!.cells.slice(16).every(({ textDirection }) => textDirection === undefined)).toBe(true);
    expect(pkg.mutations).toEqual(journal);

    (snapshot[0]!.cells[0] as { verticalAlignment?: string }).verticalAlignment = 'bottom';
    expect(table.rows[0]!.cells[0]!.verticalAlignment).toBe('top');
    expect(pkg.mutations).toEqual(journal);
  });

  it('losslessly edits one table-cell vertical alignment and rolls back atomically', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const part = pkg.requirePart('/ppt/slides/slide1.xml');
    const textBody = (bodyProperties: string, text: string): string =>
      `<a:txBody>${bodyProperties}<a:p><a:r><a:t>${text}</a:t></a:r></a:p></a:txBody>`;
    const cell = (body: string, properties: string, attributes = ''): string =>
      `<a:tc${attributes}>${body}${properties}</a:tc>`;
    const adjacentCell = cell(
      textBody('<a:bodyPr anchor="t" keep="ADJACENT"><a:spAutoFit/></a:bodyPr>', 'Adjacent'),
      '<a:tcPr anchor="b" vert="horz" keep="ADJACENT-TCPR"/>',
    );
    const cells = [
      cell(
        textBody(
          '<a:bodyPr anchor="b" custom="BODY"><a:normAutofit fontScale="85000" lnSpcReduction="20000"/><x:keep xmlns:x="urn:test">BODY</x:keep></a:bodyPr>',
          'Same mode',
        ),
        '<a:tcPr marL="100" anchor=\'ctr\' vert="vert270" custom="KEEP"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><x:keep xmlns:x="urn:test">TCPR</x:keep></a:tcPr>',
      ),
      cell(textBody('<a:bodyPr/>', 'Add top'), '<a:tcPr marR="200"/>'),
      cell(textBody('<a:bodyPr><a:noAutofit/></a:bodyPr>', 'Top to middle'), '<a:tcPr anchor="t" horzOverflow="clip"><a:noFill/></a:tcPr>'),
      cell(textBody('<a:bodyPr/>', 'Clear bottom'), '<a:tcPr anchor="b" marT="300"/>'),
      cell(textBody('<a:bodyPr/>', 'Canonicalize'), '<a:tcPr anchor="distributed" keep="UNKNOWN"/>'),
      cell(textBody('<a:bodyPr/>', 'Clear unknown'), '<a:tcPr anchor="mid" keep="CLEAR"/>'),
      adjacentCell,
      cell(
        textBody('<a:bodyPr><a:spAutoFit/></a:bodyPr>', 'Merged'),
        '<a:tcPr anchor="ctr" vert="wordArtVert" keep="MERGED"/>',
        ' hMerge="1"',
      ),
      cell(textBody('<a:bodyPr/>', 'Repeated anchor'), '<a:tcPr anchor="t" anchor="b"/>'),
      cell(textBody('<a:bodyPr/>', 'Repeated tcPr'), '<a:tcPr/><a:tcPr keep="SECOND"/>'),
      cell(textBody('<a:bodyPr/>', 'Missing tcPr'), ''),
    ].join('');
    const source = new TextDecoder().decode(part.bytes)
      .replace(
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr><a:graphic>',
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="2000"/></p:xfrm><a:graphic>',
      )
      .replace(/<a:tr>.*?<\/a:tr>/, `<a:tr>${cells}</a:tr>`);
    pkg.setPart(part.uri, source, part.contentType);
    const model = new PresentationModel(pkg);
    const table = model.slides[1]!.shapes[2] as TableModel;

    const noOpJournal = [...pkg.mutations];
    table.setCellVerticalAlignment(0, 0, 'middle');
    expect(pkg.mutations).toEqual(noOpJournal);

    table.setCellVerticalAlignment(0, 1, 'top');
    table.setCellVerticalAlignment(0, 2, 'middle');
    table.setCellVerticalAlignment(0, 3, undefined);
    table.setCellVerticalAlignment(0, 4, 'bottom');
    table.setCellVerticalAlignment(0, 5, undefined);
    table.setCellVerticalAlignment(0, 7, 'top');
    let updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain('<a:tcPr marL="100" anchor=\'ctr\' vert="vert270" custom="KEEP">');
    expect(updated).toContain('<a:tcPr marR="200" anchor="t"/>');
    expect(updated).toContain('<a:tcPr anchor="ctr" horzOverflow="clip"><a:noFill/></a:tcPr>');
    expect(updated).toContain('<a:tcPr marT="300"/>');
    expect(updated).toContain('<a:tcPr anchor="b" keep="UNKNOWN"/>');
    expect(updated).toContain('<a:tcPr keep="CLEAR"/>');
    expect(updated).toContain('<a:tc hMerge="1">');
    expect(updated).toContain('<a:tcPr anchor="t" vert="wordArtVert" keep="MERGED"/>');
    expect(updated).toContain('<a:bodyPr anchor="b" custom="BODY"><a:normAutofit fontScale="85000"');
    expect(updated).toContain(adjacentCell);

    table.setCellText(0, 2, 'Edited alignment');
    table.setCellTextDirection(0, 1, 'wordArtVert');
    table.setCellTextFit(0, 0, 'resize');
    table.setTransform({ x: inches(1) });
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain('<a:t>Edited alignment</a:t>');
    expect(updated).toContain('<a:tcPr marR="200" anchor="t" vert="wordArtVert"/>');
    expect(updated).toContain('<a:tcPr marL="100" anchor=\'ctr\' vert="vert270" custom="KEEP">');
    expect(updated).toContain('<a:bodyPr anchor="b" custom="BODY"><a:spAutoFit/>');
    expect(updated).toContain('<a:off x="914400" y="0"/>');
    expect(updated).toContain(adjacentCell);

    table.setCellVerticalAlignment(0, 2, 'bottom');
    table.setCellVerticalAlignment(0, 2, undefined);
    expect(table.rows[0]!.cells[2]!.verticalAlignment).toBeUndefined();
    expect(table.rows[0]!.cells[2]!.textFit).toBe('none');

    const beforeInvalid = pkg.requirePart(part.uri).bytes;
    const invalidJournal = [...pkg.mutations];
    for (const [row, column] of [[-1, 0], [0, -1], [0, 11], [1, 0]]) {
      expect(() => table.setCellVerticalAlignment(row!, column!, 'middle')).toThrow(RangeError);
    }
    for (const column of [8, 9, 10]) {
      expect(() => table.setCellVerticalAlignment(0, column, 'middle')).toThrow(ModelParseError);
    }
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);

    const beforeRollback = pkg.requirePart(part.uri).bytes;
    const rollbackJournal = [...pkg.mutations];
    const rollbackAlignments = table.rows[0]!.cells.map(({ verticalAlignment }) => verticalAlignment);
    expect(() =>
      pkg.transaction(() => {
        table.setCellVerticalAlignment(0, 0, 'top');
        table.setCellVerticalAlignment(0, 4, undefined);
        throw new Error('restore table cell vertical alignments');
      }),
    ).toThrow('restore table cell vertical alignments');
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(model.slides[1]!.shapes[2]).toBe(table);
    expect(table.rows[0]!.cells.map(({ verticalAlignment }) => verticalAlignment)).toEqual(
      rollbackAlignments,
    );
  });

  it('reads only exact direct single-paragraph table-cell horizontal alignments', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const part = pkg.requirePart('/ppt/slides/slide1.xml');
    const paragraph = (properties: string, text: string): string =>
      `<a:p>${properties}<a:r><a:t>${text}</a:t></a:r></a:p>`;
    const cell = (body: string, properties = '<a:tcPr/>'): string =>
      `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${body}</a:txBody>${properties}</a:tc>`;
    const cells = [
      cell(paragraph('<a:pPr algn="l"/>', 'Left')),
      cell(paragraph('<a:pPr algn="ctr"/>', 'Center')),
      cell(paragraph('<a:pPr algn="r"/>', 'Right')),
      cell(paragraph('<a:pPr algn="just"/>', 'Justify')),
      cell(paragraph('', 'Missing pPr')),
      cell(paragraph('<a:pPr/>', 'Missing algn')),
      cell(paragraph('<a:pPr algn=""/>', 'Empty')),
      cell(paragraph('<a:pPr algn="L"/>', 'Case')),
      cell(paragraph('<a:pPr algn=" ctr "/>', 'Whitespace')),
      cell(paragraph('<a:pPr algn="left"/>', 'Long form')),
      cell(paragraph('<a:pPr algn="dist"/>', 'Distributed')),
      cell(paragraph('<a:pPr algn="thaiDist"/>', 'Thai distributed')),
      cell(paragraph('<a:pPr algn="justLow"/>', 'Low justify')),
      cell(paragraph('<a:pPr algn="unknown"/>', 'Unknown')),
      cell(paragraph('<a:pPr xmlns:x="urn:test" x:algn="ctr"/>', 'Namespaced')),
      cell(paragraph('<a:pPr algn="l" algn="r"/>', 'Repeated attribute')),
      cell(paragraph('<a:pPr algn="l"/><a:pPr keep="SECOND"/>', 'Repeated pPr')),
      cell(
        paragraph('<a:pPr algn="l"/>', 'First')
        + paragraph('<a:pPr algn="r"/>', 'Second'),
      ),
      '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/></a:txBody><a:tcPr/></a:tc>',
      '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>First body</a:t></a:r></a:p></a:txBody><a:txBody><a:bodyPr/><a:p><a:r><a:t>Second body</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>',
      cell(
        paragraph('', 'Impostors'),
        '<a:tcPr algn="ctr"><x:pPr xmlns:x="urn:test" algn="r"/></a:tcPr>',
      ),
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(/<a:tr>.*?<\/a:tr>/, `<a:tr>${cells}</a:tr>`),
      part.contentType,
    );
    const table = new PresentationModel(pkg).slides[1]!.shapes[2] as TableModel;
    const journal = [...pkg.mutations];

    const snapshot = table.rows;
    expect(snapshot[0]!.cells.slice(0, 4).map(({ horizontalAlignment }) =>
      horizontalAlignment)).toEqual(['left', 'center', 'right', 'justify']);
    expect(snapshot[0]!.cells.slice(4).every(({ horizontalAlignment }) =>
      horizontalAlignment === undefined)).toBe(true);
    expect(pkg.mutations).toEqual(journal);

    (snapshot[0]!.cells[0] as { horizontalAlignment?: string }).horizontalAlignment = 'right';
    expect(table.rows[0]!.cells[0]!.horizontalAlignment).toBe('left');
    expect(pkg.mutations).toEqual(journal);
  });

  it('losslessly edits one table-cell horizontal alignment and rolls back atomically', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const part = pkg.requirePart('/ppt/slides/slide1.xml');
    const textBody = (paragraphs: string, bodyProperties = '<a:bodyPr/>'): string =>
      `<a:txBody>${bodyProperties}<a:lstStyle/>${paragraphs}</a:txBody>`;
    const paragraph = (properties: string, content: string): string =>
      `<a:p>${properties}${content}</a:p>`;
    const run = (text: string): string => `<a:r><a:t>${text}</a:t></a:r>`;
    const cell = (body: string, properties: string, attributes = ''): string =>
      `<a:tc${attributes}>${body}${properties}</a:tc>`;
    const adjacentCell = cell(
      textBody(paragraph('<a:pPr algn="r" keep="NEIGHBOR"/>', run('Neighbor'))),
      '<a:tcPr anchor="b" vert="horz" keep="NEIGHBOR-TCPR"/>',
    );
    const cells = [
      cell(
        textBody(
          paragraph(
            "<a:pPr algn='ctr' marL=\"111\" custom=\"SAME\"><a:buNone/><x:keep xmlns:x=\"urn:test\">SAME</x:keep></a:pPr>",
            run('Same center'),
          ),
          '<a:bodyPr custom="SAME-BODY"><a:normAutofit fontScale="85000"/></a:bodyPr>',
        ),
        '<a:tcPr marL="100" anchor="ctr" vert="vert270" keep="SAME-TCPR"/>',
      ),
      cell(
        textBody(paragraph('', run('Missing pPr'))),
        '<a:tcPr marR="200"/>',
      ),
      cell(
        textBody(paragraph('<a:pPr/>', run('Self closing pPr'))),
        '<a:tcPr anchor="t"/>',
      ),
      cell(
        textBody(paragraph(
          '<a:pPr algn="l" marL="222" custom="EXPANDED"><a:spcBef><a:spcPts val="300"/></a:spcBef><a:buNone/><x:keep xmlns:x="urn:test">PPR</x:keep></a:pPr>',
          '<a:r><a:rPr lang="fr-FR"/><a:t>Expanded</a:t></a:r><a:br/><a:fld id="{TEST}"><a:rPr lang="en-US"/><a:t>Field</a:t></a:fld><a:endParaRPr lang="ja-JP"/>',
        )),
        '<a:tcPr horzOverflow="clip"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:tcPr>',
      ),
      cell(
        textBody(paragraph('<a:pPr algn="just" keep="CLEAR"/>', run('Clear justify'))),
        '<a:tcPr marT="300"/>',
      ),
      cell(
        textBody(paragraph('<a:pPr algn="dist" keep="UNKNOWN"/>', run('Canonicalize unknown'))),
        '<a:tcPr keep="UNKNOWN-TCPR"/>',
      ),
      cell(
        textBody(paragraph(
          '<a:pPr xmlns:x="urn:test" x:algn="ctr" keep="NAMESPACED"/>',
          run('Namespaced only'),
        )),
        '<a:tcPr/>',
      ),
      cell(
        textBody(paragraph('<a:pPr keep="MERGED"/>', run('Merged placeholder'))),
        '<a:tcPr vert="wordArtVert" keep="MERGED-TCPR"/>',
        ' hMerge="1"',
      ),
      adjacentCell,
      cell(
        textBody(paragraph('<a:pPr algn="l" algn="r"/>', run('Repeated algn'))),
        '<a:tcPr/>',
      ),
      cell(
        textBody(paragraph('<a:pPr/><a:pPr keep="SECOND"/>', run('Repeated pPr'))),
        '<a:tcPr/>',
      ),
      cell(
        textBody(
          paragraph('<a:pPr algn="l"/>', run('First paragraph'))
          + paragraph('<a:pPr algn="r"/>', run('Second paragraph')),
        ),
        '<a:tcPr/>',
      ),
      cell(textBody(''), '<a:tcPr/>'),
      '<a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>First body</a:t></a:r></a:p></a:txBody><a:txBody><a:bodyPr/><a:p><a:r><a:t>Second body</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>',
    ].join('');
    const source = new TextDecoder().decode(part.bytes)
      .replace(
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr><a:graphic>',
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="2000"/></p:xfrm><a:graphic>',
      )
      .replace(/<a:tr>.*?<\/a:tr>/, `<a:tr>${cells}</a:tr>`);
    pkg.setPart(part.uri, source, part.contentType);
    const model = new PresentationModel(pkg);
    const table = model.slides[1]!.shapes[2] as TableModel;

    const noOpBefore = pkg.requirePart(part.uri).bytes.slice();
    const noOpJournal = [...pkg.mutations];
    table.setCellHorizontalAlignment(0, 0, 'center');
    table.setCellHorizontalAlignment(0, 6, undefined);
    expect(pkg.requirePart(part.uri).bytes).toEqual(noOpBefore);
    expect(pkg.mutations).toEqual(noOpJournal);

    table.setCellHorizontalAlignment(0, 1, 'left');
    table.setCellHorizontalAlignment(0, 2, 'right');
    table.setCellHorizontalAlignment(0, 3, 'justify');
    table.setCellHorizontalAlignment(0, 4, undefined);
    table.setCellHorizontalAlignment(0, 5, 'center');
    table.setCellHorizontalAlignment(0, 7, 'left');
    let updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain("<a:pPr algn='ctr' marL=\"111\" custom=\"SAME\">");
    expect(updated).toContain('<a:p><a:pPr algn="l"/><a:r><a:t>Missing pPr</a:t></a:r></a:p>');
    expect(updated).toContain('<a:pPr algn="r"/>');
    expect(updated).toContain('<a:pPr algn="just" marL="222" custom="EXPANDED">');
    expect(updated).toContain('<a:spcBef><a:spcPts val="300"/></a:spcBef><a:buNone/><x:keep xmlns:x="urn:test">PPR</x:keep>');
    expect(updated).toContain('<a:r><a:rPr lang="fr-FR"/><a:t>Expanded</a:t></a:r><a:br/><a:fld id="{TEST}"><a:rPr lang="en-US"/><a:t>Field</a:t></a:fld><a:endParaRPr lang="ja-JP"/>');
    expect(updated).toContain('<a:pPr keep="CLEAR"/>');
    expect(updated).toContain('<a:pPr algn="ctr" keep="UNKNOWN"/>');
    expect(updated).toContain('<a:pPr xmlns:x="urn:test" x:algn="ctr" keep="NAMESPACED"/>');
    expect(updated).toContain('<a:tc hMerge="1">');
    expect(updated).toContain('<a:pPr keep="MERGED" algn="l"/>');
    expect(updated).toContain(adjacentCell);
    expect(table.rows[0]!.cells.slice(0, 9).map(({ horizontalAlignment }) =>
      horizontalAlignment)).toEqual([
      'center',
      'left',
      'right',
      'justify',
      undefined,
      'center',
      undefined,
      'left',
      'right',
    ]);

    const expandedParagraph = table.rows[0]!.cells[3]!.richText[0]!;
    table.setCellRichText(0, 3, [{
      ...expandedParagraph,
      runs: [{ text: 'Edited expanded' }],
    }]);
    table.setCellTextDirection(0, 1, 'wordArtVert');
    table.setCellTextFit(0, 0, 'resize');
    table.setCellVerticalAlignment(0, 2, 'bottom');
    table.setTransform({ x: inches(1) });
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain('>Edited expanded</a:t>');
    expect(updated).toContain('<a:pPr algn="just" marL="222" custom="EXPANDED">');
    expect(updated).toContain('<a:tcPr marR="200" vert="wordArtVert"/>');
    expect(updated).toContain('<a:bodyPr custom="SAME-BODY"><a:spAutoFit/>');
    expect(updated).toContain('<a:tcPr anchor="b"/>');
    expect(updated).toContain('<a:off x="914400" y="0"/>');
    expect(updated).toContain(adjacentCell);

    const beforeInvalid = pkg.requirePart(part.uri).bytes.slice();
    const invalidJournal = [...pkg.mutations];
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
      expect(() => table.setCellHorizontalAlignment(0, 0, value as never)).toThrow(TypeError);
    }
    for (const [row, column] of [
      [-1, 0],
      [0, -1],
      [0.5, 0],
      [0, 0.5],
      [Number.NaN, 0],
      [0, Number.NaN],
      [Number.POSITIVE_INFINITY, 0],
      [0, Number.NEGATIVE_INFINITY],
      [1, 0],
      [0, 14],
    ]) {
      expect(() => table.setCellHorizontalAlignment(row!, column!, 'center')).toThrow(RangeError);
    }
    for (const column of [9, 10, 11, 12, 13]) {
      expect(() => table.setCellHorizontalAlignment(0, column, 'center')).toThrow(ModelParseError);
    }
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeInvalid);
    expect(pkg.mutations).toEqual(invalidJournal);

    const beforeRollback = pkg.requirePart(part.uri).bytes.slice();
    const rollbackJournal = [...pkg.mutations];
    const rollbackAlignments = table.rows[0]!.cells.map(({ horizontalAlignment }) =>
      horizontalAlignment);
    expect(() => pkg.transaction(() => {
      table.setCellHorizontalAlignment(0, 0, 'left');
      table.setCellHorizontalAlignment(0, 5, undefined);
      throw new Error('restore table cell horizontal alignments');
    })).toThrow('restore table cell horizontal alignments');
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(model.slides[1]!.shapes[2]).toBe(table);
    expect(table.rows[0]!.cells.map(({ horizontalAlignment }) =>
      horizontalAlignment)).toEqual(rollbackAlignments);
  });

  it('reads paragraph and soft breaks, then preserves the first paragraph style on plain-text overwrite', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const richText = new TextDecoder()
      .decode(part.bytes)
      .replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="fr-FR"/><a:t>First</a:t></a:r><a:br/><a:r><a:t>soft</a:t></a:r></a:p><x:keep xmlns:x="urn:test">OPAQUE</x:keep><a:p><a:r><a:t>Second</a:t></a:r></a:p>',
      );
    pkg.setPart(part.uri, richText, part.contentType);
    const shape = slide.shapes[0] as ShapeModel;

    expect(shape.text).toBe('First\nsoft\nSecond');
    expect(slide.title.text).toBe('First\nsoft\nSecond');
    shape.text = ' Updated \r\n\rEnd';
    expect(shape.text).toBe(' Updated \n\nEnd');
    expect(slide.title.text).toBe(' Updated \n\nEnd');
    expect(slide.shapes[0]).toBe(shape);

    const updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated.match(/<a:pPr algn="ctr"\/>/g)).toHaveLength(3);
    expect(updated.match(/<a:rPr lang="fr-FR"\/>/g)).toHaveLength(3);
    expect(updated).not.toContain('<a:br/>');
    expect(updated).toContain('<x:keep xmlns:x="urn:test">OPAQUE</x:keep>');
    expect(updated).toContain('<a:t xml:space="preserve"> Updated </a:t>');

    expect(() =>
      pkg.transaction(() => {
        shape.text = 'Rollback';
        throw new Error('restore rich text');
      }),
    ).toThrow('restore rich text');
    expect(shape.text).toBe(' Updated \n\nEnd');

    const selfClosingPkg = await OpcPackage.open(await modelFixture());
    const selfClosingModel = new PresentationModel(selfClosingPkg);
    const selfClosingSlide = selfClosingModel.slides[1]!;
    const selfClosingPart = selfClosingPkg.requirePart(selfClosingSlide.partUri);
    selfClosingPkg.setPart(
      selfClosingPart.uri,
      new TextDecoder().decode(selfClosingPart.bytes).replace('<a:t>First title</a:t>', '<a:t/>'),
      selfClosingPart.contentType,
    );
    const selfClosingShape = selfClosingSlide.shapes[0] as ShapeModel;
    selfClosingShape.text = 'Expanded';
    expect(selfClosingShape.text).toBe('Expanded');
    expect(new TextDecoder().decode(selfClosingPkg.requirePart(selfClosingPart.uri).bytes)).toContain(
      '<a:t xml:space="preserve">Expanded</a:t>',
    );
  });

  it('reads fields and repeated soft breaks, then preserves paragraph and text-body metadata on rich replacement', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const richText = new TextDecoder()
      .decode(part.bytes)
      .replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        '<p:txBodyMeta xmlns:p="urn:test">KEEP</p:txBodyMeta><a:p><a:pPr algn="r"><a:buNone/></a:pPr><a:r><a:t>First</a:t></a:r><a:br/><a:br/><a:fld id="{TEST}"><a:rPr lang="en-US" sz="1250" b="0" i="1"><a:solidFill><a:srgbClr val="00aa11"/></a:solidFill><a:latin typeface="Field Font"/></a:rPr><a:t>Field</a:t></a:fld><a:br/><a:endParaRPr lang="fr-FR"/></a:p>',
      );
    pkg.setPart(part.uri, richText, part.contentType);
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.richText).toEqual([
      {
        align: 'right',
        runs: [
          { text: 'First' },
          { text: '', softBreakBefore: true },
          {
            text: 'Field',
            softBreakBefore: true,
            style: {
              fontFamily: 'Field Font',
              fontSize: 12.5,
              lang: 'en-US',
              bold: false,
              italic: true,
              color: { kind: 'srgb', value: '00AA11' },
            },
          },
          { text: '', softBreakBefore: true },
        ],
      },
    ]);
    expect(shape.text).toBe('First\n\nField\n');
    expect(pkg.mutations).toEqual(journal);

    shape.richText = [
      {
        runs: [{ text: 'One', style: { bold: true } }, { text: 'Two', softBreakBefore: true }],
        align: 'center',
      },
      { runs: [] },
    ];
    expect(shape.text).toBe('One\nTwo\n');
    expect(slide.shapes[0]).toBe(shape);
    const updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated.match(/<a:pPr algn="ctr"><a:buNone\/><\/a:pPr>/g)).toHaveLength(1);
    expect(updated.match(/<a:pPr><a:buNone\/><\/a:pPr>/g)).toHaveLength(1);
    expect(updated.match(/<a:endParaRPr lang="fr-FR"\/>/g)).toHaveLength(2);
    expect(updated).toContain('<p:txBodyMeta xmlns:p="urn:test">KEEP</p:txBodyMeta>');
    expect(updated).not.toContain('<a:fld');
  });

  it('reads and losslessly replaces strict direct text-box margins', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const marginBody = [
      '<a:bodyPr wrap="square" lIns="127000" tIns="-6350"',
      ' rIns="2147483647" bIns="1e3" custom="KEEP">',
      '<a:normAutofit fontScale="90000"/>',
      '<x:keep xmlns:x="urn:test">KEEP</x:keep>',
      '</a:bodyPr>',
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace('<p:txBody><a:p>', `<p:txBody>${marginBody}<a:p>`),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.textMargins).toEqual({
      left: 10,
      top: -0.5,
      right: 2_147_483_647 / 12_700,
    });
    expect(pkg.mutations).toEqual(journal);
    const detached = shape.textMargins as { left?: number };
    detached.left = 99;
    expect(shape.textMargins?.left).toBe(10);

    shape.textMargins = { top: 4, left: 8 };
    expect(shape.textMargins).toEqual({ left: 8, top: 4 });
    expect(slide.shapes[0]).toBe(shape);
    let updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain(
      '<a:bodyPr wrap="square" lIns="101600" tIns="50800" custom="KEEP">',
    );
    expect(updated).not.toMatch(/\s[rb]Ins=/);
    expect(updated).toContain('<a:normAutofit fontScale="90000"/>');
    expect(updated).toContain('<x:keep xmlns:x="urn:test">KEEP</x:keep>');
    expect(shape.text).toBe('First title');

    shape.text = 'Plain replacement';
    shape.richText = [{ runs: [{ text: 'Rich replacement', style: { bold: true } }] }];
    shape.setTransform({ x: inches(3) });
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain(
      '<a:bodyPr wrap="square" lIns="101600" tIns="50800" custom="KEEP">',
    );

    const beforeRollback = pkg.requirePart(part.uri).bytes.slice();
    expect(() =>
      pkg.transaction(() => {
        shape.textMargins = 10;
        throw new Error('restore text margins');
      }),
    ).toThrow('restore text margins');
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeRollback);
    expect(shape.textMargins).toEqual({ left: 8, top: 4 });
    expect(slide.shapes[0]).toBe(shape);

    shape.textMargins = undefined;
    expect(shape.textMargins).toBeUndefined();
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).not.toMatch(/\s(?:lIns|tIns|rIns|bIns)=/);
    expect(updated).toContain('<a:bodyPr wrap="square" custom="KEEP">');
    expect(updated).toContain('<a:normAutofit fontScale="90000"/>');
    expect(updated).toContain('<x:keep xmlns:x="urn:test">KEEP</x:keep>');
  });

  it('ignores malformed direct text-box margin attributes independently', async () => {
    for (const malformed of ['100.5', '1e3', '', '+100', '2147483648', '-2147483649']) {
      const pkg = await OpcPackage.open(await modelFixture());
      const model = new PresentationModel(pkg);
      const slide = model.slides[1]!;
      const part = pkg.requirePart(slide.partUri);
      pkg.setPart(
        part.uri,
        new TextDecoder().decode(part.bytes).replace(
          '<p:txBody><a:p>',
          `<p:txBody><a:bodyPr lIns="${malformed}" rIns="12700"/><a:p>`,
        ),
        part.contentType,
      );
      const shape = slide.shapes[0] as ShapeModel;
      const journal = [...pkg.mutations];

      expect(shape.textMargins).toEqual({ right: 1 });
      expect(pkg.mutations).toEqual(journal);
    }
  });

  it('reads and losslessly replaces direct text-box vertical alignment', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const bodyProperties = [
      '<a:bodyPr wrap="square" lIns="127000" anchor="t" anchorCtr="1" custom="KEEP">',
      '<a:normAutofit fontScale="90000"/>',
      '<x:keep xmlns:x="urn:test">KEEP</x:keep>',
      '</a:bodyPr>',
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<p:txBody><a:p>',
        `<p:txBody>${bodyProperties}<a:p>`,
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.verticalAlignment).toBe('top');
    expect(shape.verticalAlignment).toBe('top');
    expect(pkg.mutations).toEqual(journal);
    expect(shape.textMargins).toEqual({ left: 10 });

    shape.verticalAlignment = 'middle';
    expect(shape.verticalAlignment).toBe('middle');
    let updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain('lIns="127000" anchor="ctr" anchorCtr="1" custom="KEEP"');

    shape.verticalAlignment = 'bottom';
    expect(shape.verticalAlignment).toBe('bottom');
    shape.text = 'Plain replacement';
    shape.richText = [{ runs: [{ text: 'Rich replacement', style: { bold: true } }] }];
    shape.textMargins = { top: 4, left: 8 };
    shape.setTransform({ x: inches(3) });
    expect(shape.verticalAlignment).toBe('bottom');
    expect(slide.shapes[0]).toBe(shape);
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain('anchor="b" anchorCtr="1" custom="KEEP"');
    expect(updated).toContain('<a:normAutofit fontScale="90000"/>');
    expect(updated).toContain('<x:keep xmlns:x="urn:test">KEEP</x:keep>');

    const beforeRollback = pkg.requirePart(part.uri).bytes.slice();
    expect(() =>
      pkg.transaction(() => {
        shape.verticalAlignment = 'top';
        throw new Error('restore vertical alignment');
      }),
    ).toThrow('restore vertical alignment');
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeRollback);
    expect(shape.verticalAlignment).toBe('bottom');
    expect(slide.shapes[0]).toBe(shape);

    shape.verticalAlignment = undefined;
    expect(shape.verticalAlignment).toBeUndefined();
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).not.toMatch(/\sanchor=/);
    expect(updated).toContain('anchorCtr="1" custom="KEEP"');
    expect(updated).toContain('lIns="101600"');
    expect(updated).toContain('tIns="50800"');
    expect(updated).toContain('<a:normAutofit fontScale="90000"/>');
    expect(updated).toContain('<x:keep xmlns:x="urn:test">KEEP</x:keep>');
  });

  it('reads only supported direct text-box vertical-alignment tokens', async () => {
    const cases: readonly [string | undefined, string | undefined][] = [
      ['t', 'top'],
      ['ctr', 'middle'],
      ['b', 'bottom'],
      [undefined, undefined],
      ['just', undefined],
      ['dist', undefined],
      ['', undefined],
      ['T', undefined],
      [' middle ', undefined],
      ['unknown', undefined],
    ];
    for (const [token, expected] of cases) {
      const pkg = await OpcPackage.open(await modelFixture());
      const model = new PresentationModel(pkg);
      const slide = model.slides[1]!;
      const part = pkg.requirePart(slide.partUri);
      const anchor = token === undefined ? '' : ` anchor="${token}"`;
      pkg.setPart(
        part.uri,
        new TextDecoder().decode(part.bytes).replace(
          '<p:txBody><a:p>',
          `<p:txBody><a:bodyPr${anchor}/><a:p>`,
        ),
        part.contentType,
      );
      const shape = slide.shapes[0] as ShapeModel;
      const journal = [...pkg.mutations];

      expect(shape.verticalAlignment).toBe(expected);
      expect(pkg.mutations).toEqual(journal);
      if (token === 'just' || token === 'dist') {
        shape.textMargins = 2;
        shape.text = 'Plain replacement';
        shape.richText = [{ runs: [{ text: 'Rich replacement' }] }];
        shape.setTransform({ x: inches(2) });
        expect(new TextDecoder().decode(pkg.requirePart(part.uri).bytes)).toContain(`anchor="${token}"`);
      }
    }
  });

  it('reads and losslessly replaces direct text-box wrapping', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const bodyProperties = [
      `<a:bodyPr wrap='none' lIns="127000" anchor="b" vert="vert" custom="KEEP">`,
      '<a:normAutofit fontScale="90000"/>',
      '<x:keep xmlns:x="urn:test">KEEP</x:keep>',
      '</a:bodyPr>',
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<p:txBody><a:p>',
        `<p:txBody>${bodyProperties}<a:p>`,
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.textWrap).toBe(false);
    expect(shape.textWrap).toBe(false);
    expect(pkg.mutations).toEqual(journal);
    expect(shape.textMargins).toEqual({ left: 10 });
    expect(shape.verticalAlignment).toBe('bottom');

    shape.textWrap = true;
    expect(shape.textWrap).toBe(true);
    let updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain(`wrap='square'`);
    expect(updated).toContain('anchor="b" vert="vert" custom="KEEP"');

    shape.textWrap = false;
    expect(shape.textWrap).toBe(false);
    shape.text = 'Plain replacement';
    shape.richText = [{ runs: [{ text: 'Rich replacement', style: { bold: true } }] }];
    shape.textMargins = { top: 4, left: 8 };
    shape.verticalAlignment = 'top';
    shape.setTransform({ x: inches(3) });
    expect(shape.textWrap).toBe(false);
    expect(slide.shapes[0]).toBe(shape);
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain(`wrap='none'`);
    expect(updated).toContain('anchor="t" vert="vert" custom="KEEP"');
    expect(updated).toContain('lIns="101600"');
    expect(updated).toContain('tIns="50800"');
    expect(updated).toContain('<a:normAutofit fontScale="90000"/>');
    expect(updated).toContain('<x:keep xmlns:x="urn:test">KEEP</x:keep>');

    const beforeRollback = pkg.requirePart(part.uri).bytes.slice();
    const rollbackJournal = [...pkg.mutations];
    expect(() =>
      pkg.transaction(() => {
        shape.textWrap = true;
        throw new Error('restore text wrapping');
      }),
    ).toThrow('restore text wrapping');
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(shape.textWrap).toBe(false);
    expect(slide.shapes[0]).toBe(shape);

    shape.textWrap = undefined;
    expect(shape.textWrap).toBeUndefined();
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).not.toMatch(/\swrap=/);
    expect(updated).toContain('anchor="t" vert="vert" custom="KEEP"');
    expect(updated).toContain('<a:normAutofit fontScale="90000"/>');
    expect(updated).toContain('<x:keep xmlns:x="urn:test">KEEP</x:keep>');
  });

  it('reads only supported direct text-box wrapping tokens', async () => {
    const cases: readonly [string | undefined, boolean | undefined][] = [
      ['square', true],
      ['none', false],
      [undefined, undefined],
      ['', undefined],
      ['Square', undefined],
      [' square ', undefined],
      ['false', undefined],
      ['tight', undefined],
      ['unknown', undefined],
    ];
    for (const [token, expected] of cases) {
      const pkg = await OpcPackage.open(await modelFixture());
      const model = new PresentationModel(pkg);
      const slide = model.slides[1]!;
      const part = pkg.requirePart(slide.partUri);
      const wrap = token === undefined ? '' : ` wrap="${token}"`;
      pkg.setPart(
        part.uri,
        new TextDecoder().decode(part.bytes).replace(
          '<p:txBody><a:p>',
          `<p:txBody><a:bodyPr${wrap}/><a:p>`,
        ),
        part.contentType,
      );
      const shape = slide.shapes[0] as ShapeModel;
      const journal = [...pkg.mutations];

      expect(shape.textWrap).toBe(expected);
      expect(pkg.mutations).toEqual(journal);
      if (token === 'tight') {
        shape.text = 'Plain replacement';
        shape.richText = [{ runs: [{ text: 'Rich replacement' }] }];
        shape.textMargins = 2;
        shape.verticalAlignment = 'bottom';
        shape.setTransform({ x: inches(2) });
        expect(new TextDecoder().decode(pkg.requirePart(part.uri).bytes)).toContain('wrap="tight"');
      }
    }
  });

  it('reads and losslessly replaces direct text-box direction', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const bodyProperties = [
      `<a:bodyPr wrap="none" lIns="127000" anchor="b" vert='vert270' rtlCol="1" custom="KEEP">`,
      '<a:normAutofit fontScale="90000"/>',
      '<x:keep xmlns:x="urn:test">KEEP</x:keep>',
      '</a:bodyPr>',
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<p:txBody><a:p>',
        `<p:txBody>${bodyProperties}<a:p>`,
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.textDirection).toBe('vert270');
    expect(shape.textDirection).toBe('vert270');
    expect(pkg.mutations).toEqual(journal);
    expect(shape.textMargins).toEqual({ left: 10 });
    expect(shape.verticalAlignment).toBe('bottom');
    expect(shape.textWrap).toBe(false);

    shape.textDirection = 'wordArtVert';
    expect(shape.textDirection).toBe('wordArtVert');
    let updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain(`vert='wordArtVert'`);
    expect(updated).toContain('rtlCol="1" custom="KEEP"');

    shape.textDirection = 'eaVert';
    expect(shape.textDirection).toBe('eaVert');
    shape.text = 'Plain replacement';
    shape.richText = [{ runs: [{ text: 'Rich replacement', style: { bold: true } }] }];
    shape.textMargins = { top: 4, left: 8 };
    shape.verticalAlignment = 'top';
    shape.textWrap = true;
    shape.setTransform({ x: inches(3) });
    expect(shape.textDirection).toBe('eaVert');
    expect(slide.shapes[0]).toBe(shape);
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain(`vert='eaVert'`);
    expect(updated).toContain('wrap="square"');
    expect(updated).toContain('anchor="t"');
    expect(updated).toContain('rtlCol="1" custom="KEEP"');
    expect(updated).toContain('<a:normAutofit fontScale="90000"/>');
    expect(updated).toContain('<x:keep xmlns:x="urn:test">KEEP</x:keep>');

    const beforeRollback = pkg.requirePart(part.uri).bytes.slice();
    const rollbackJournal = [...pkg.mutations];
    expect(() =>
      pkg.transaction(() => {
        shape.textDirection = 'horz';
        throw new Error('restore text direction');
      }),
    ).toThrow('restore text direction');
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(shape.textDirection).toBe('eaVert');
    expect(slide.shapes[0]).toBe(shape);

    shape.textDirection = undefined;
    expect(shape.textDirection).toBeUndefined();
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).not.toMatch(/\svert=/);
    expect(updated).toContain('wrap="square"');
    expect(updated).toContain('anchor="t"');
    expect(updated).toContain('rtlCol="1" custom="KEEP"');
    expect(updated).toContain('<a:normAutofit fontScale="90000"/>');
    expect(updated).toContain('<x:keep xmlns:x="urn:test">KEEP</x:keep>');
  });

  it('reads only supported direct text-box direction tokens', async () => {
    const directions = [
      'eaVert',
      'horz',
      'mongolianVert',
      'vert',
      'vert270',
      'wordArtVert',
      'wordArtVertRtl',
    ] as const;
    const cases: ReadonlyArray<readonly [string | undefined, string | undefined]> = [
      ...directions.map((direction) => [direction, direction] as const),
      [undefined, undefined],
      ['', undefined],
      ['Vert', undefined],
      [' vert ', undefined],
      ['vertical', undefined],
      ['unknown', undefined],
    ];
    for (const [token, expected] of cases) {
      const pkg = await OpcPackage.open(await modelFixture());
      const model = new PresentationModel(pkg);
      const slide = model.slides[1]!;
      const part = pkg.requirePart(slide.partUri);
      const vert = token === undefined ? '' : ` vert="${token}"`;
      pkg.setPart(
        part.uri,
        new TextDecoder().decode(part.bytes).replace(
          '<p:txBody><a:p>',
          `<p:txBody><a:bodyPr${vert}/><a:p>`,
        ),
        part.contentType,
      );
      const shape = slide.shapes[0] as ShapeModel;
      const journal = [...pkg.mutations];

      expect(shape.textDirection).toBe(expected);
      expect(pkg.mutations).toEqual(journal);
      if (token === 'vertical') {
        shape.text = 'Plain replacement';
        shape.richText = [{ runs: [{ text: 'Rich replacement' }] }];
        shape.textMargins = 2;
        shape.verticalAlignment = 'bottom';
        shape.textWrap = false;
        shape.setTransform({ x: inches(2) });
        expect(new TextDecoder().decode(pkg.requirePart(part.uri).bytes)).toContain('vert="vertical"');
      }
    }
  });

  it('reads and losslessly replaces direct text-box fit', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const bodyProperties = [
      `<a:bodyPr wrap="none" lIns="127000" anchor="b" vert="vert270" custom="KEEP">`,
      '<a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp>',
      `<a:normAutofit fontScale='85000' lnSpcReduction="20000"/>`,
      '<a:scene3d><a:camera prst="orthographicFront"/></a:scene3d>',
      '<x:keep xmlns:x="urn:test">KEEP</x:keep>',
      '</a:bodyPr>',
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<p:txBody><a:p>',
        `<p:txBody>${bodyProperties}<a:p>`,
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const beforeSameMode = pkg.requirePart(part.uri).bytes.slice();
    const journal = [...pkg.mutations];

    expect(shape.textFit).toBe('shrink');
    expect(shape.textFit).toBe('shrink');
    expect(pkg.mutations).toEqual(journal);
    shape.textFit = 'shrink';
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeSameMode);
    expect(pkg.mutations).toEqual(journal);

    shape.textFit = 'resize';
    expect(shape.textFit).toBe('resize');
    let updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain(
      '<a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp>'
      + '<a:spAutoFit/>'
      + '<a:scene3d><a:camera prst="orthographicFront"/></a:scene3d>',
    );
    expect(updated).not.toContain('normAutofit');
    expect(updated).toContain('<x:keep xmlns:x="urn:test">KEEP</x:keep>');

    shape.text = 'Plain replacement';
    shape.richText = [{ runs: [{ text: 'Rich replacement', style: { bold: true } }] }];
    shape.textMargins = { top: 4, left: 8 };
    shape.verticalAlignment = 'top';
    shape.textWrap = true;
    shape.textDirection = 'wordArtVert';
    shape.setTransform({ x: inches(3) });
    expect(shape.textFit).toBe('resize');
    expect(slide.shapes[0]).toBe(shape);
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain('<a:spAutoFit/>');
    expect(updated).toContain('wrap="square"');
    expect(updated).toContain('anchor="t"');
    expect(updated).toContain('vert="wordArtVert"');
    expect(updated).toContain('custom="KEEP"');
    expect(updated).toContain('<x:keep xmlns:x="urn:test">KEEP</x:keep>');

    const beforeRollback = pkg.requirePart(part.uri).bytes.slice();
    const rollbackJournal = [...pkg.mutations];
    expect(() =>
      pkg.transaction(() => {
        shape.textFit = 'shrink';
        throw new Error('restore text fit');
      }),
    ).toThrow('restore text fit');
    expect(pkg.requirePart(part.uri).bytes).toEqual(beforeRollback);
    expect(pkg.mutations).toEqual(rollbackJournal);
    expect(shape.textFit).toBe('resize');

    shape.textFit = undefined;
    expect(shape.textFit).toBeUndefined();
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).not.toMatch(/<(?:a:)?(?:noAutofit|normAutofit|spAutoFit)\b/);
    expect(updated).toContain(
      '<a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp>'
      + '<a:scene3d><a:camera prst="orthographicFront"/></a:scene3d>',
    );

    shape.textFit = 'shrink';
    expect(shape.textFit).toBe('shrink');
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain(
      '<a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp>'
      + '<a:normAutofit/>'
      + '<a:scene3d><a:camera prst="orthographicFront"/></a:scene3d>',
    );
    shape.textFit = 'none';
    expect(shape.textFit).toBeUndefined();
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).not.toMatch(/<(?:a:)?(?:noAutofit|normAutofit|spAutoFit)\b/);
    expect(updated).toContain('<x:keep xmlns:x="urn:test">KEEP</x:keep>');
  });

  it('reads only a unique direct text-box fit choice', async () => {
    const cases: ReadonlyArray<readonly [string, string | undefined]> = [
      ['', undefined],
      ['<a:noAutofit/>', 'none'],
      ['<a:normAutofit/>', 'shrink'],
      ['<a:spAutoFit/>', 'resize'],
      ['<a:NormAutofit/>', undefined],
      ['<x:normAutofit xmlns:x="urn:test"/>', undefined],
      ['<x:keep xmlns:x="urn:test"><a:normAutofit/></x:keep>', undefined],
      ['<a:normAutofit/><a:normAutofit/>', undefined],
      ['<a:noAutofit/><a:spAutoFit/>', undefined],
      ['<x:unknown xmlns:x="urn:test"/>', undefined],
    ];
    for (const [children, expected] of cases) {
      const pkg = await OpcPackage.open(await modelFixture());
      const model = new PresentationModel(pkg);
      const slide = model.slides[1]!;
      const part = pkg.requirePart(slide.partUri);
      const bodyProperties = children === ''
        ? '<a:bodyPr/>'
        : `<a:bodyPr>${children}</a:bodyPr>`;
      pkg.setPart(
        part.uri,
        new TextDecoder().decode(part.bytes).replace(
          '<p:txBody><a:p>',
          `<p:txBody>${bodyProperties}<a:p>`,
        ),
        part.contentType,
      );
      const shape = slide.shapes[0] as ShapeModel;
      const journal = [...pkg.mutations];

      expect(shape.textFit).toBe(expected);
      expect(pkg.mutations).toEqual(journal);
      if (children === '<a:noAutofit/>') {
        shape.text = 'Plain replacement';
        shape.richText = [{ runs: [{ text: 'Rich replacement' }] }];
        shape.textMargins = 2;
        shape.verticalAlignment = 'bottom';
        shape.textWrap = false;
        shape.textDirection = 'vert';
        shape.setTransform({ x: inches(2) });
        expect(new TextDecoder().decode(pkg.requirePart(part.uri).bytes)).toContain('<a:noAutofit/>');
      }
    }
  });

  it('reads strict local underline values and preserves their XML during plain text edits', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const underlineText = [
      '<a:p>',
      '<a:r><a:rPr u="words"><a:uFill><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></a:uFill></a:rPr><a:t>Words</a:t></a:r>',
      '<a:r><a:rPr u="none"><a:uFill><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:uFill></a:rPr><a:t>None</a:t></a:r>',
      '<a:r><a:rPr u="mystery"><a:uFill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:uFill></a:rPr><a:t>Unknown</a:t></a:r>',
      '<a:r><a:rPr u="dbl"><a:uFill><a:solidFill><a:srgbClr val="red"/></a:solidFill></a:uFill></a:rPr><a:t>Malformed</a:t></a:r>',
      '<a:r><a:rPr><a:uFill><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:uFill></a:rPr><a:t>Orphan</a:t></a:r>',
      '<a:r><a:rPr u="wavyHeavy"><a:uFill><a:solidFill><a:schemeClr val="tx2"/></a:solidFill></a:uFill></a:rPr><a:t>Scheme</a:t></a:r>',
      '<a:r><a:rPr u="dash"><a:uFill><a:solidFill><a:schemeClr val="accent1"><a:tint val="50000"/></a:schemeClr></a:solidFill></a:uFill></a:rPr><a:t>Transformed</a:t></a:r>',
      '</a:p>',
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        underlineText,
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.richText[0]!.runs.map(({ style }) => style?.underline)).toEqual([
      { style: 'words', color: { kind: 'srgb', value: 'FF0000' } },
      false,
      undefined,
      { style: 'dbl' },
      undefined,
      { style: 'wavyHeavy', color: { kind: 'scheme', value: 'tx2' } },
      { style: 'dash' },
    ]);
    expect(pkg.mutations).toEqual(journal);

    const snapshot = shape.richText as unknown as Array<{
      runs: Array<{ style?: { underline?: { style?: string; color?: { value: string } } } }>;
    }>;
    snapshot[0]!.runs[0]!.style!.underline!.style = 'dbl';
    snapshot[0]!.runs[0]!.style!.underline!.color!.value = '000000';
    expect(shape.richText[0]!.runs[0]!.style!.underline).toEqual({
      style: 'words',
      color: { kind: 'srgb', value: 'FF0000' },
    });

    shape.text = 'First replacement\nSecond replacement';
    expect(shape.richText.map((paragraph) => paragraph.runs[0]!.style!.underline)).toEqual([
      { style: 'words', color: { kind: 'srgb', value: 'FF0000' } },
      { style: 'words', color: { kind: 'srgb', value: 'FF0000' } },
    ]);
    const updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated.match(/<a:rPr u="words">/g)).toHaveLength(2);
    expect(updated.match(/<a:uFill><a:solidFill><a:srgbClr val="ff0000"\/><\/a:solidFill><\/a:uFill>/g))
      .toHaveLength(2);
  });

  it('reads strict local strike values and preserves their XML during plain text edits', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const strikeText = [
      '<a:p>',
      '<a:r><a:rPr strike="sngStrike"/><a:t>Single</a:t></a:r>',
      '<a:r><a:rPr strike="dblStrike" u="sng"/><a:t>Double</a:t></a:r>',
      '<a:r><a:rPr strike="noStrike"/><a:t>None</a:t></a:r>',
      '<a:r><a:rPr strike="tripleStrike" b="1"/><a:t>Unknown</a:t></a:r>',
      '<a:r><a:rPr/><a:t>Absent</a:t></a:r>',
      '</a:p>',
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        strikeText,
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.richText[0]!.runs.map(({ style }) => style?.strike)).toEqual([
      'sngStrike',
      'dblStrike',
      false,
      undefined,
      undefined,
    ]);
    expect(shape.richText[0]!.runs[1]!.style!.underline).toEqual({ style: 'sng' });
    expect(shape.richText[0]!.runs[3]!.style!.bold).toBe(true);
    expect(pkg.mutations).toEqual(journal);

    shape.text = 'First replacement\nSecond replacement';
    expect(shape.richText.map((paragraph) => paragraph.runs[0]!.style!.strike)).toEqual([
      'sngStrike',
      'sngStrike',
    ]);
    const updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated.match(/<a:rPr strike="sngStrike"\/>/g)).toHaveLength(2);
  });

  it('reads strict local highlight colors and preserves their XML during plain text edits', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const highlightText = [
      '<a:p>',
      '<a:r><a:rPr><a:highlight><a:srgbClr val="ffff00"/></a:highlight></a:rPr><a:t>Yellow</a:t></a:r>',
      '<a:r><a:rPr><a:highlight><a:schemeClr val="accent2"/></a:highlight></a:rPr><a:t>Theme</a:t></a:r>',
      '<a:r><a:rPr><a:highlight><a:srgbClr val="FF0000"/></a:highlight><a:highlight><a:srgbClr val="00FF00"/></a:highlight></a:rPr><a:t>Repeated</a:t></a:r>',
      '<a:r><a:rPr><a:highlight><a:prstClr val="yellow"/></a:highlight></a:rPr><a:t>Unsupported</a:t></a:r>',
      '<a:r><a:rPr><a:highlight><a:srgbClr val="yellow"/></a:highlight></a:rPr><a:t>Malformed</a:t></a:r>',
      '<a:r><a:rPr strike="sngStrike"><a:highlight><a:schemeClr val="accent1"><a:tint val="50000"/></a:schemeClr></a:highlight></a:rPr><a:t>Transformed</a:t></a:r>',
      '<a:r><a:rPr><a:highlight><a:srgbClr val="112233"/><a:schemeClr val="tx1"/></a:highlight></a:rPr><a:t>Multiple</a:t></a:r>',
      '<a:r><a:rPr/><a:t>Absent</a:t></a:r>',
      '</a:p>',
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        highlightText,
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.richText[0]!.runs.map(({ style }) => style?.highlight)).toEqual([
      { kind: 'srgb', value: 'FFFF00' },
      { kind: 'scheme', value: 'accent2' },
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    expect(shape.richText[0]!.runs[5]!.style!.strike).toBe('sngStrike');
    expect(pkg.mutations).toEqual(journal);

    const snapshot = shape.richText as unknown as Array<{
      runs: Array<{ style?: { highlight?: { value: string } } }>;
    }>;
    snapshot[0]!.runs[0]!.style!.highlight!.value = '000000';
    expect(shape.richText[0]!.runs[0]!.style!.highlight).toEqual({ kind: 'srgb', value: 'FFFF00' });

    shape.text = 'First replacement\nSecond replacement';
    expect(shape.richText.map((paragraph) => paragraph.runs[0]!.style!.highlight)).toEqual([
      { kind: 'srgb', value: 'FFFF00' },
      { kind: 'srgb', value: 'FFFF00' },
    ]);
    const updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated.match(/<a:highlight><a:srgbClr val="ffff00"\/><\/a:highlight>/g)).toHaveLength(2);
  });

  it('reads strict local text outlines and preserves their XML during plain text edits', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const outlineText = [
      '<a:p>',
      '<a:r><a:rPr><a:ln w="19050"><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></a:ln></a:rPr><a:t>Red</a:t></a:r>',
      '<a:r><a:rPr><a:ln w="0"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln></a:rPr><a:t>Hairline</a:t></a:r>',
      '<a:r><a:rPr><a:ln w="20116800"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill></a:ln></a:rPr><a:t>Maximum</a:t></a:r>',
      '<a:r><a:rPr><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></a:rPr><a:t>Repeated</a:t></a:r>',
      '<a:r><a:rPr><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:rPr><a:t>Missing width</a:t></a:r>',
      '<a:r><a:rPr><a:ln w="-1"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:rPr><a:t>Negative</a:t></a:r>',
      '<a:r><a:rPr><a:ln w="20116801"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:rPr><a:t>Too wide</a:t></a:r>',
      '<a:r><a:rPr><a:ln w="1.5"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:rPr><a:t>Decimal</a:t></a:r>',
      '<a:r><a:rPr><a:ln w="12700" cap="round"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></a:rPr><a:t>Cap</a:t></a:r>',
      '<a:r><a:rPr strike="dblStrike"><a:ln w="12700"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:prstDash val="dash"/></a:ln></a:rPr><a:t>Dash</a:t></a:r>',
      '<a:r><a:rPr><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:ln></a:rPr><a:t>Transform</a:t></a:r>',
      '<a:r><a:rPr><a:ln w="12700"><a:noFill/></a:ln></a:rPr><a:t>Unknown fill</a:t></a:r>',
      '<a:r><a:rPr><a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln></a:rPr><a:t>Multiple fills</a:t></a:r>',
      '<a:r><a:rPr/><a:t>Absent</a:t></a:r>',
      '</a:p>',
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        outlineText,
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.richText[0]!.runs.map(({ style }) => style?.outline)).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, size: 1.5 },
      { color: { kind: 'scheme', value: 'accent1' }, size: 0 },
      { color: { kind: 'scheme', value: 'tx2' }, size: 1584 },
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    expect(shape.richText[0]!.runs[9]!.style!.strike).toBe('dblStrike');
    expect(pkg.mutations).toEqual(journal);

    const snapshot = shape.richText as unknown as Array<{
      runs: Array<{ style?: { outline?: { color: { value: string }; size: number } } }>;
    }>;
    snapshot[0]!.runs[0]!.style!.outline!.color.value = '000000';
    snapshot[0]!.runs[0]!.style!.outline!.size = 3;
    expect(shape.richText[0]!.runs[0]!.style!.outline).toEqual({
      color: { kind: 'srgb', value: 'FF0000' },
      size: 1.5,
    });

    shape.text = 'First replacement\nSecond replacement';
    expect(shape.richText.map((paragraph) => paragraph.runs[0]!.style!.outline)).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, size: 1.5 },
      { color: { kind: 'srgb', value: 'FF0000' }, size: 1.5 },
    ]);
    const updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated.match(/<a:ln w="19050"><a:solidFill><a:srgbClr val="ff0000"\/><\/a:solidFill><\/a:ln>/g))
      .toHaveLength(2);
  });

  it('reads strict local text glows and preserves their XML during plain text edits', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const glowText = [
      '<a:p>',
      '<a:r><a:rPr><a:effectLst><a:glow rad="101600"><a:srgbClr val="ff0000"><a:alpha val="50000"/></a:srgbClr></a:glow></a:effectLst></a:rPr><a:t>Red</a:t></a:r>',
      '<a:r><a:rPr><a:effectLst><a:glow rad="31750"><a:schemeClr val="accent1"/></a:glow></a:effectLst></a:rPr><a:t>Theme</a:t></a:r>',
      '<a:r><a:rPr><a:effectLst><a:glow rad="0"><a:srgbClr val="00FF00"><a:alpha val="0"/></a:srgbClr></a:glow></a:effectLst></a:rPr><a:t>Zero</a:t></a:r>',
      '<a:r><a:rPr><a:effectLst><a:glow rad="27273042316900"><a:schemeClr val="accent2"><a:alpha val="100000"/></a:schemeClr></a:glow></a:effectLst></a:rPr><a:t>Maximum</a:t></a:r>',
      '<a:r><a:rPr><a:effectLst><a:glow rad="12700"><a:srgbClr val="0000FF"><a:alpha val="75000"/></a:srgbClr></a:glow><a:outerShdw blurRad="0"/></a:effectLst></a:rPr><a:t>Sibling effect</a:t></a:r>',
      '<a:r><a:rPr><a:effectLst><a:glow rad="12700"><a:srgbClr val="FF0000"/></a:glow></a:effectLst><a:effectLst><a:glow rad="12700"><a:srgbClr val="00FF00"/></a:glow></a:effectLst></a:rPr><a:t>Repeated list</a:t></a:r>',
      '<a:r><a:rPr><a:effectDag/><a:effectLst><a:glow rad="12700"><a:srgbClr val="FF0000"/></a:glow></a:effectLst></a:rPr><a:t>Effect DAG</a:t></a:r>',
      '<a:r><a:rPr><a:effectLst><a:glow rad="12700"><a:srgbClr val="FF0000"/></a:glow><a:glow rad="25400"><a:srgbClr val="00FF00"/></a:glow></a:effectLst></a:rPr><a:t>Repeated glow</a:t></a:r>',
      '<a:r><a:rPr><a:effectLst><a:glow><a:srgbClr val="FF0000"/></a:glow></a:effectLst></a:rPr><a:t>Missing radius</a:t></a:r>',
      '<a:r><a:rPr><a:effectLst><a:glow rad="-1"><a:srgbClr val="FF0000"/></a:glow></a:effectLst></a:rPr><a:t>Negative</a:t></a:r>',
      '<a:r><a:rPr><a:effectLst><a:glow rad="27273042316901"><a:srgbClr val="FF0000"/></a:glow></a:effectLst></a:rPr><a:t>Too large</a:t></a:r>',
      '<a:r><a:rPr><a:effectLst><a:glow rad="1.5"><a:srgbClr val="FF0000"/></a:glow></a:effectLst></a:rPr><a:t>Decimal</a:t></a:r>',
      '<a:r><a:rPr><a:effectLst><a:glow rad="12700" custom="1"><a:srgbClr val="FF0000"/></a:glow></a:effectLst></a:rPr><a:t>Extra attribute</a:t></a:r>',
      '<a:r><a:rPr><a:effectLst><a:glow rad="12700"><a:prstClr val="red"/></a:glow></a:effectLst></a:rPr><a:t>Unsupported color</a:t></a:r>',
      '<a:r><a:rPr strike="sngStrike"><a:effectLst><a:glow rad="12700"><a:schemeClr val="accent1"><a:tint val="50000"/></a:schemeClr></a:glow></a:effectLst></a:rPr><a:t>Transform</a:t></a:r>',
      '<a:r><a:rPr><a:effectLst><a:glow rad="12700"><a:srgbClr val="FF0000"><a:alpha val="50000"/><a:alpha val="75000"/></a:srgbClr></a:glow></a:effectLst></a:rPr><a:t>Repeated alpha</a:t></a:r>',
      '<a:r><a:rPr><a:effectLst><a:glow rad="12700"><a:srgbClr val="FF0000"><a:alpha val="100001"/></a:srgbClr></a:glow></a:effectLst></a:rPr><a:t>High alpha</a:t></a:r>',
      '<a:r><a:rPr><a:effectLst><a:glow rad="12700"><a:srgbClr val="FF0000"><a:alpha val="1.5"/></a:srgbClr></a:glow></a:effectLst></a:rPr><a:t>Decimal alpha</a:t></a:r>',
      '<a:r><a:rPr/><a:t>Absent</a:t></a:r>',
      '</a:p>',
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        glowText,
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.richText[0]!.runs.map(({ style }) => style?.glow)).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, opacity: 0.5, size: 8 },
      { color: { kind: 'scheme', value: 'accent1' }, opacity: 1, size: 2.5 },
      { color: { kind: 'srgb', value: '00FF00' }, opacity: 0, size: 0 },
      { color: { kind: 'scheme', value: 'accent2' }, opacity: 1, size: 2_147_483_647 },
      { color: { kind: 'srgb', value: '0000FF' }, opacity: 0.75, size: 1 },
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    expect(shape.richText[0]!.runs[14]!.style!.strike).toBe('sngStrike');
    expect(pkg.mutations).toEqual(journal);

    const snapshot = shape.richText as unknown as Array<{
      runs: Array<{ style?: { glow?: { color?: { value: string }; opacity: number; size: number } } }>;
    }>;
    snapshot[0]!.runs[0]!.style!.glow!.color!.value = '000000';
    snapshot[0]!.runs[0]!.style!.glow!.opacity = 1;
    snapshot[0]!.runs[0]!.style!.glow!.size = 3;
    expect(shape.richText[0]!.runs[0]!.style!.glow).toEqual({
      color: { kind: 'srgb', value: 'FF0000' },
      opacity: 0.5,
      size: 8,
    });

    shape.text = 'First replacement\nSecond replacement';
    expect(shape.richText.map((paragraph) => paragraph.runs[0]!.style!.glow)).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, opacity: 0.5, size: 8 },
      { color: { kind: 'srgb', value: 'FF0000' }, opacity: 0.5, size: 8 },
    ]);
    const updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated.match(/<a:effectLst><a:glow rad="101600"><a:srgbClr val="ff0000"><a:alpha val="50000"\/><\/a:srgbClr><\/a:glow><\/a:effectLst>/g))
      .toHaveLength(2);
  });

  it('reads strict local main-text transparency without mutating malformed color transforms', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const transparencyText = [
      '<a:p>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="0"/></a:srgbClr></a:solidFill></a:rPr><a:t>Invisible</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"><a:alpha val="1"/></a:schemeClr></a:solidFill></a:rPr><a:t>Almost invisible</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="00FF00"><a:alpha val="49445"/></a:srgbClr></a:solidFill></a:rPr><a:t>Fractional</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:schemeClr val="tx1"><a:alpha val="75000"/></a:schemeClr></a:solidFill></a:rPr><a:t>Quarter</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="0000FF"><a:alpha val="100000"/></a:srgbClr></a:solidFill></a:rPr><a:t>Opaque</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>No alpha</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha/></a:srgbClr></a:solidFill></a:rPr><a:t>Missing value</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val=""/></a:srgbClr></a:solidFill></a:rPr><a:t>Empty</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="1.5"/></a:srgbClr></a:solidFill></a:rPr><a:t>Decimal</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="1e3"/></a:srgbClr></a:solidFill></a:rPr><a:t>Scientific</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="-1"/></a:srgbClr></a:solidFill></a:rPr><a:t>Negative</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="100001"/></a:srgbClr></a:solidFill></a:rPr><a:t>Too high</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/><a:alpha val="75000"/></a:srgbClr></a:solidFill></a:rPr><a:t>Repeated alpha</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"><a:alphaMod val="50000"/></a:srgbClr></a:solidFill></a:rPr><a:t>Alpha mod</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"><a:alphaOff val="50000"/></a:srgbClr></a:solidFill></a:rPr><a:t>Alpha offset</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/><a:tint val="50000"/></a:srgbClr></a:solidFill></a:rPr><a:t>Mixed transform</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000" custom="KEEP"/></a:srgbClr></a:solidFill></a:rPr><a:t>Alpha attribute</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"><a:ext/></a:alpha></a:srgbClr></a:solidFill></a:rPr><a:t>Alpha child</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000" custom="KEEP"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:rPr><a:t>Color attribute</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill><a:solidFill><a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr></a:solidFill></a:rPr><a:t>Repeated fill</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr><a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr></a:solidFill></a:rPr><a:t>Repeated color</a:t></a:r>',
      '<a:r><a:rPr><a:solidFill><a:srgbClr val="red"><a:alpha val="50000"/></a:srgbClr></a:solidFill></a:rPr><a:t>Invalid sRGB</a:t></a:r>',
      '<a:r><a:rPr strike="sngStrike"><a:solidFill><a:schemeClr val="unknown"><a:alpha val="50000"/></a:schemeClr></a:solidFill></a:rPr><a:t>Unknown scheme</a:t></a:r>',
      '</a:p>',
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        transparencyText,
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    const transparencies = shape.richText[0]!.runs.map(({ style }) => style?.transparency);
    expect(transparencies.slice(0, 5)).toEqual([100, 99.999, 50.555, 25, 0]);
    expect(transparencies.slice(5)).toEqual(Array(18).fill(undefined));
    expect(shape.richText[0]!.runs[22]!.style!.strike).toBe('sngStrike');
    expect(pkg.mutations).toEqual(journal);
  });

  it('preserves main-text alpha and unrelated run XML during plain text edits', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        '<a:p><a:r><a:rPr strike="sngStrike"><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="75000"/></a:srgbClr></a:solidFill><a:effectLst><a:outerShdw blurRad="0"/></a:effectLst></a:rPr><a:t>Original</a:t></a:r></a:p>',
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;

    shape.text = 'First\nSecond';

    expect(shape.richText.map((paragraph) => paragraph.runs[0]!.style!.transparency)).toEqual([25, 25]);
    const updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated.match(/<a:srgbClr val="FF0000"><a:alpha val="75000"\/><\/a:srgbClr>/g)).toHaveLength(2);
    expect(updated.match(/<a:effectLst><a:outerShdw blurRad="0"\/><\/a:effectLst>/g)).toHaveLength(2);
  });

  it('reads strict local text baselines and preserves their XML during plain text edits', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const baselineText = [
      '<a:p>',
      '<a:r><a:rPr baseline="30000"/><a:t>Super</a:t></a:r>',
      '<a:r><a:rPr baseline="-40000"/><a:t>Sub</a:t></a:r>',
      '<a:r><a:rPr baseline="0"/><a:t>Normal</a:t></a:r>',
      '<a:r><a:rPr baseline="12346"/><a:t>Custom</a:t></a:r>',
      '<a:r><a:rPr baseline="-2147483648"/><a:t>Minimum</a:t></a:r>',
      '<a:r><a:rPr baseline="2147483647"/><a:t>Maximum</a:t></a:r>',
      '<a:r><a:rPr baseline="1.5"/><a:t>Decimal</a:t></a:r>',
      '<a:r><a:rPr baseline="1e3" strike="dblStrike"/><a:t>Exponent</a:t></a:r>',
      '<a:r><a:rPr baseline="2147483648"/><a:t>Too high</a:t></a:r>',
      '<a:r><a:rPr baseline="-2147483649"/><a:t>Too low</a:t></a:r>',
      '<a:r><a:rPr baseline=""/><a:t>Empty</a:t></a:r>',
      '<a:r><a:rPr/><a:t>Absent</a:t></a:r>',
      '</a:p>',
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        baselineText,
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.richText[0]!.runs.map(({ style }) => style?.baseline)).toEqual([
      'superscript',
      'subscript',
      0,
      12.346,
      -2_147_483.648,
      2_147_483.647,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    expect(shape.richText[0]!.runs[7]!.style!.strike).toBe('dblStrike');
    expect(pkg.mutations).toEqual(journal);

    shape.text = 'First replacement\nSecond replacement';
    expect(shape.richText.map((paragraph) => paragraph.runs[0]!.style!.baseline)).toEqual([
      'superscript',
      'superscript',
    ]);
    const updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated.match(/<a:rPr baseline="30000"\/>/g)).toHaveLength(2);
  });

  it('reads strict local character spacing and preserves its XML during plain text edits', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const spacingText = [
      '<a:p>',
      '<a:r><a:rPr spc="250" kern="1200"/><a:t>Expanded</a:t></a:r>',
      '<a:r><a:rPr spc="-125" kern="0"/><a:t>Condensed</a:t></a:r>',
      '<a:r><a:rPr spc="0" kern="1200"/><a:t>Normal</a:t></a:r>',
      '<a:r><a:rPr spc="-2147483648"/><a:t>Minimum</a:t></a:r>',
      '<a:r><a:rPr spc="2147483647"/><a:t>Maximum</a:t></a:r>',
      '<a:r><a:rPr spc="1.5"/><a:t>Decimal</a:t></a:r>',
      '<a:r><a:rPr spc="1e3" strike="sngStrike"/><a:t>Exponent</a:t></a:r>',
      '<a:r><a:rPr spc=""/><a:t>Empty</a:t></a:r>',
      '<a:r><a:rPr spc="2147483648"/><a:t>Too high</a:t></a:r>',
      '<a:r><a:rPr spc="-2147483649"/><a:t>Too low</a:t></a:r>',
      '<a:r><a:rPr kern="0"/><a:t>Kern only</a:t></a:r>',
      '<a:r><a:rPr/><a:t>Absent</a:t></a:r>',
      '</a:p>',
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        spacingText,
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.richText[0]!.runs.map(({ style }) => style?.characterSpacing)).toEqual([
      2.5,
      -1.25,
      0,
      -21_474_836.48,
      21_474_836.47,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    expect(shape.richText[0]!.runs[6]!.style!.strike).toBe('sngStrike');
    expect(pkg.mutations).toEqual(journal);

    shape.text = 'First replacement\nSecond replacement';
    expect(shape.richText.map((paragraph) => paragraph.runs[0]!.style!.characterSpacing)).toEqual([2.5, 2.5]);
    const updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated.match(/<a:rPr spc="250" kern="1200"\/>/g)).toHaveLength(2);
  });

  it('reads direct run languages and preserves language XML during plain text edits', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const languageText = [
      '<a:p>',
      '<a:r><a:rPr lang="fr-CA" altLang="it-IT"/><a:t>French</a:t></a:r>',
      '<a:r><a:rPr lang="x-private"/><a:t>Private</a:t></a:r>',
      '<a:r><a:rPr lang="" strike="sngStrike"/><a:t>Empty</a:t></a:r>',
      '<a:r><a:rPr altLang="ja-JP" b="1"/><a:t>Alternate only</a:t></a:r>',
      '<a:endParaRPr lang="zh-CN" altLang="ja-JP"/>',
      '</a:p>',
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        languageText,
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.richText[0]!.runs.map(({ style }) => style?.lang)).toEqual([
      'fr-CA',
      'x-private',
      undefined,
      undefined,
    ]);
    expect(shape.richText[0]!.runs[2]!.style!.strike).toBe('sngStrike');
    expect(shape.richText[0]!.runs[3]!.style!.bold).toBe(true);
    expect(pkg.mutations).toEqual(journal);

    shape.text = 'First replacement\nSecond replacement';
    expect(shape.richText.map((paragraph) => paragraph.runs[0]!.style!.lang)).toEqual([
      'fr-CA',
      'fr-CA',
    ]);
    const updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated.match(/<a:rPr lang="fr-CA" altLang="it-IT"\/>/g)).toHaveLength(2);
    expect(updated.match(/<a:endParaRPr lang="zh-CN" altLang="ja-JP"\/>/g)).toHaveLength(2);
  });

  it('reads strict direct paragraph RTL and preserves its XML during plain text edits', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const values = ['1', 'true', 'on', '0', 'false', 'off', undefined, '', 'yes'];
    const rtlText = values.map((rtl, index) => [
      '<a:p>',
      `<a:pPr${rtl === undefined ? '' : ` rtl="${rtl}"`} custom="P${index}"><x:keep xmlns:x="urn:test"/></a:pPr>`,
      `<a:r><a:t>Paragraph ${index}</a:t></a:r>`,
      '</a:p>',
    ].join('')).join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        rtlText,
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.richText.map(({ rtl }) => rtl)).toEqual([
      true,
      true,
      true,
      false,
      false,
      false,
      undefined,
      undefined,
      undefined,
    ]);
    expect(pkg.mutations).toEqual(journal);

    shape.text = 'First replacement\nSecond replacement';
    expect(shape.richText.map(({ rtl }) => rtl)).toEqual([true, true]);
    const updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated.match(/<a:pPr rtl="1" custom="P0"><x:keep xmlns:x="urn:test"\/><\/a:pPr>/g))
      .toHaveLength(2);
  });

  it('reads strict non-list paragraph left margins and preserves them during plain text edits', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const marginText = [
      '<a:p><a:pPr marL="0"/><a:r><a:t>Zero</a:t></a:r></a:p>',
      '<a:p><a:pPr marL="152400"/><a:r><a:t>Twelve</a:t></a:r></a:p>',
      '<a:p><a:pPr marL="1"/><a:r><a:t>One EMU</a:t></a:r></a:p>',
      '<a:p><a:pPr/><a:r><a:t>Missing</a:t></a:r></a:p>',
      '<a:p><a:pPr marL=""/><a:r><a:t>Empty</a:t></a:r></a:p>',
      '<a:p><a:pPr marL="1.5"/><a:r><a:t>Fractional</a:t></a:r></a:p>',
      '<a:p><a:pPr marL="-1"/><a:r><a:t>Negative</a:t></a:r></a:p>',
      '<a:p><a:pPr marL="51206401"/><a:r><a:t>Too large</a:t></a:r></a:p>',
      '<a:p><a:pPr marL="yes"/><a:r><a:t>Unknown</a:t></a:r></a:p>',
      '<a:p><a:pPr marL="228600"><a:buNone/></a:pPr><a:r><a:t>No bullet</a:t></a:r></a:p>',
      '<a:p><a:pPr marL="228600"><a:buChar char="•"/></a:pPr><a:r><a:t>Bullet</a:t></a:r></a:p>',
      '<a:p><a:pPr marL="228600"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>Number</a:t></a:r></a:p>',
      '<a:p><a:pPr marL="228600"><a:buBlip><a:blip r:embed="rId1"/></a:buBlip></a:pPr><a:r><a:t>Picture</a:t></a:r></a:p>',
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        marginText,
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.richText.map(({ marginLeft }) => marginLeft)).toEqual([
      0,
      12,
      1 / 12700,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      18,
      undefined,
      undefined,
      undefined,
    ]);
    expect(pkg.mutations).toEqual(journal);

    const plainPkg = await OpcPackage.open(await modelFixture());
    const plainModel = new PresentationModel(plainPkg);
    const plainSlide = plainModel.slides[1]!;
    const plainPart = plainPkg.requirePart(plainSlide.partUri);
    plainPkg.setPart(
      plainPart.uri,
      new TextDecoder().decode(plainPart.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        '<a:p><a:pPr marL="152400" custom="KEEP"><x:keep xmlns:x="urn:test"/></a:pPr><a:r><a:t>Original</a:t></a:r></a:p>',
      ),
      plainPart.contentType,
    );
    const plainShape = plainSlide.shapes[0] as ShapeModel;
    plainShape.text = 'First\nSecond';
    expect(plainShape.richText.map(({ marginLeft }) => marginLeft)).toEqual([12, 12]);
    const updated = new TextDecoder().decode(plainPkg.requirePart(plainPart.uri).bytes);
    expect(updated.match(/<a:pPr marL="152400" custom="KEEP"><x:keep xmlns:x="urn:test"\/><\/a:pPr>/g))
      .toHaveLength(2);
  });

  it('reads strict paragraph right margins on ordinary and list paragraphs and preserves them during plain text edits', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const marginText = [
      '<a:p><a:pPr marR="0"/><a:r><a:t>Zero</a:t></a:r></a:p>',
      '<a:p><a:pPr marR="152400"/><a:r><a:t>Twelve</a:t></a:r></a:p>',
      '<a:p><a:pPr marR="1"/><a:r><a:t>One EMU</a:t></a:r></a:p>',
      '<a:p><a:pPr/><a:r><a:t>Missing</a:t></a:r></a:p>',
      '<a:p><a:pPr marR=""/><a:r><a:t>Empty</a:t></a:r></a:p>',
      '<a:p><a:pPr marR="1.5"/><a:r><a:t>Fractional</a:t></a:r></a:p>',
      '<a:p><a:pPr marR="-1"/><a:r><a:t>Negative</a:t></a:r></a:p>',
      '<a:p><a:pPr marR="51206401"/><a:r><a:t>Too large</a:t></a:r></a:p>',
      '<a:p><a:pPr marR="yes"/><a:r><a:t>Unknown</a:t></a:r></a:p>',
      '<a:p><a:pPr marR="228600"><a:buNone/></a:pPr><a:r><a:t>No bullet</a:t></a:r></a:p>',
      '<a:p><a:pPr marR="228600"><a:buChar char="•"/></a:pPr><a:r><a:t>Bullet</a:t></a:r></a:p>',
      '<a:p><a:pPr marR="228600"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>Number</a:t></a:r></a:p>',
      '<a:p><a:pPr marR="228600"><a:buBlip><a:blip r:embed="rId1"/></a:buBlip></a:pPr><a:r><a:t>Picture</a:t></a:r></a:p>',
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        marginText,
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.richText.map(({ marginRight }) => marginRight)).toEqual([
      0,
      12,
      1 / 12700,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      18,
      18,
      18,
      18,
    ]);
    expect(pkg.mutations).toEqual(journal);

    const plainPkg = await OpcPackage.open(await modelFixture());
    const plainModel = new PresentationModel(plainPkg);
    const plainSlide = plainModel.slides[1]!;
    const plainPart = plainPkg.requirePart(plainSlide.partUri);
    plainPkg.setPart(
      plainPart.uri,
      new TextDecoder().decode(plainPart.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        '<a:p><a:pPr marR="152400" custom="KEEP"><x:keep xmlns:x="urn:test"/></a:pPr><a:r><a:t>Original</a:t></a:r></a:p>',
      ),
      plainPart.contentType,
    );
    const plainShape = plainSlide.shapes[0] as ShapeModel;
    plainShape.text = 'First\nSecond';
    expect(plainShape.richText.map(({ marginRight }) => marginRight)).toEqual([12, 12]);
    const updated = new TextDecoder().decode(plainPkg.requirePart(plainPart.uri).bytes);
    expect(updated.match(/<a:pPr marR="152400" custom="KEEP"><x:keep xmlns:x="urn:test"\/><\/a:pPr>/g))
      .toHaveLength(2);
  });

  it('reads strict signed ordinary paragraph indents and preserves them during plain text edits', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const indentText = [
      '<a:p><a:pPr indent="0"/><a:r><a:t>Zero</a:t></a:r></a:p>',
      '<a:p><a:pPr indent="152400"/><a:r><a:t>First line</a:t></a:r></a:p>',
      '<a:p><a:pPr indent="-228600"/><a:r><a:t>Hanging</a:t></a:r></a:p>',
      '<a:p><a:pPr indent="1"/><a:r><a:t>One EMU</a:t></a:r></a:p>',
      '<a:p><a:pPr/><a:r><a:t>Missing</a:t></a:r></a:p>',
      '<a:p><a:pPr indent=""/><a:r><a:t>Empty</a:t></a:r></a:p>',
      '<a:p><a:pPr indent="1.5"/><a:r><a:t>Fractional</a:t></a:r></a:p>',
      '<a:p><a:pPr indent="51206401"/><a:r><a:t>Too positive</a:t></a:r></a:p>',
      '<a:p><a:pPr indent="-51206401"/><a:r><a:t>Too negative</a:t></a:r></a:p>',
      '<a:p><a:pPr indent="yes"/><a:r><a:t>Unknown</a:t></a:r></a:p>',
      '<a:p><a:pPr indent="228600"><a:buNone/></a:pPr><a:r><a:t>No bullet</a:t></a:r></a:p>',
      '<a:p><a:pPr indent="228600"><a:buChar char="•"/></a:pPr><a:r><a:t>Bullet</a:t></a:r></a:p>',
      '<a:p><a:pPr indent="228600"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:t>Number</a:t></a:r></a:p>',
      '<a:p><a:pPr indent="228600"><a:buBlip><a:blip r:embed="rId1"/></a:buBlip></a:pPr><a:r><a:t>Picture</a:t></a:r></a:p>',
    ].join('');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        indentText,
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.richText.map(({ indent }) => indent)).toEqual([
      0,
      12,
      -18,
      1 / 12700,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      18,
      undefined,
      undefined,
      undefined,
    ]);
    expect(pkg.mutations).toEqual(journal);

    const plainPkg = await OpcPackage.open(await modelFixture());
    const plainModel = new PresentationModel(plainPkg);
    const plainSlide = plainModel.slides[1]!;
    const plainPart = plainPkg.requirePart(plainSlide.partUri);
    plainPkg.setPart(
      plainPart.uri,
      new TextDecoder().decode(plainPart.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        '<a:p><a:pPr marL="304800" marR="152400" indent="-228600" custom="KEEP"><x:keep xmlns:x="urn:test"/></a:pPr><a:r><a:t>Original</a:t></a:r></a:p>',
      ),
      plainPart.contentType,
    );
    const plainShape = plainSlide.shapes[0] as ShapeModel;
    plainShape.text = 'First\nSecond';
    expect(plainShape.richText.map(({ indent }) => indent)).toEqual([-18, -18]);
    const updated = new TextDecoder().decode(plainPkg.requirePart(plainPart.uri).bytes);
    expect(updated.match(/<a:pPr marL="304800" marR="152400" indent="-228600" custom="KEEP"><x:keep xmlns:x="urn:test"\/><\/a:pPr>/g))
      .toHaveLength(2);
  });

  it('updates alignment without rebuilding other paragraph properties', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        '<a:p><a:pPr algn="l" marL="111" custom="FIRST"><a:buNone/><x:first xmlns:x="urn:test"/></a:pPr><a:r><a:t>First</a:t></a:r></a:p><a:p><a:pPr algn="r" marL="222" custom="SECOND"><a:spcBef><a:spcPts val="300"/></a:spcBef></a:pPr><a:r><a:t>Second</a:t></a:r></a:p><a:p><a:pPr algn="dist" marL="333" custom="THIRD"><x:third xmlns:x="urn:test"/></a:pPr><a:r><a:t>Third</a:t></a:r></a:p><a:p><a:pPr/><a:r><a:t>Fourth</a:t></a:r></a:p>',
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;

    expect(shape.richText.map(({ align, marginLeft }) => ({ align, marginLeft }))).toEqual([
      { align: 'left', marginLeft: 111 / 12700 },
      { align: 'right', marginLeft: 222 / 12700 },
      { align: undefined, marginLeft: 333 / 12700 },
      { align: undefined, marginLeft: undefined },
    ]);
    shape.richText = [
      { runs: [{ text: 'First' }], align: 'center', marginLeft: 111 / 12700 },
      { runs: [{ text: 'Second' }], align: 'justify', marginLeft: 222 / 12700 },
      { runs: [{ text: 'Third' }], marginLeft: 333 / 12700 },
      { runs: [{ text: 'Fourth' }], align: 'right' },
      { runs: [], align: 'left', marginLeft: 111 / 12700 },
    ];

    const updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain('<a:pPr algn="ctr" marL="111" custom="FIRST"><a:buNone/><x:first xmlns:x="urn:test"/></a:pPr>');
    expect(updated).toContain('<a:pPr algn="just" marL="222" custom="SECOND"><a:buNone/></a:pPr>');
    expect(updated).toContain('<a:pPr marL="333" custom="THIRD"><x:third xmlns:x="urn:test"/><a:buNone/></a:pPr>');
    expect(updated).toContain('<a:pPr algn="r"><a:buNone/></a:pPr>');
    expect(updated).toContain('<a:pPr algn="l" marL="111" custom="FIRST"><a:buNone/><x:first xmlns:x="urn:test"/></a:pPr>');
    expect(shape.richText.map(({ align }) => align)).toEqual(['center', 'justify', undefined, 'right', 'left']);
  });

  it('reads and replaces bullet choices without rebuilding unrelated paragraph properties', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        '<a:p><a:pPr algn="r" marL="228600" indent="-228600" custom="FIRST"><a:spcBef><a:spcPts val="300"/></a:spcBef><a:buClr><a:srgbClr val="FF0000"/></a:buClr><a:buSzPts val="1400"/><a:buFont typeface="Wingdings"/><a:buChar char="◆"/><a:tabLst><a:tab pos="500000" algn="l"/></a:tabLst><a:defRPr sz="1600"/><a:extLst><a:ext uri="urn:test"><x:keep xmlns:x="urn:test">KEEP</x:keep></a:ext></a:extLst></a:pPr><a:r><a:t>Bullet</a:t></a:r></a:p><a:p><a:pPr marL="304800" indent="-100000" custom="SECOND"><a:spcAft><a:spcPts val="400"/></a:spcAft><a:buSzPct val="100000"/><a:buFont typeface="+mj-lt"/><a:buAutoNum type="romanLcParenR" startAt="4"/></a:pPr><a:r><a:t>Number</a:t></a:r></a:p><a:p><a:pPr><a:buAutoNum type="hebrew2Minus" startAt="2"/></a:pPr><a:r><a:t>Unknown number</a:t></a:r></a:p><a:p><a:pPr><a:buBlip><a:blip r:embed="rId1"/></a:buBlip></a:pPr><a:r><a:t>Picture bullet</a:t></a:r></a:p>',
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.richText.map(({ bullet }) => bullet)).toEqual([
      { kind: 'bullet', character: '◆', indent: 18 },
      { kind: 'number', style: 'romanLcParenR', startAt: 4, indent: 24 },
      undefined,
      undefined,
    ]);
    expect(pkg.mutations).toEqual(journal);

    shape.richText = [
      {
        runs: [{ text: 'Numbered' }],
        align: 'left',
        bullet: { kind: 'number', style: 'alphaUcPeriod', startAt: 2, indent: 20 },
        tabStops: [{ position: 500000 / 914400, alignment: 'left' }],
      },
      { runs: [{ text: 'No bullet' }], align: 'center', marginLeft: 24 },
      { runs: [{ text: 'Custom' }], bullet: { kind: 'bullet', character: '▶', indent: 19 } },
      { runs: [{ text: 'Cleared picture' }] },
    ];

    const updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    const firstStart = updated.indexOf('<a:pPr algn="l" marL="254000" indent="-254000" custom="FIRST">');
    const firstEnd = updated.indexOf('</a:pPr>', firstStart);
    const firstProperties = updated.slice(firstStart, firstEnd);
    expect(firstStart).toBeGreaterThan(-1);
    expect(firstProperties.indexOf('<a:spcBef>')).toBeLessThan(firstProperties.indexOf('<a:buSzPct'));
    expect(firstProperties.indexOf('<a:buSzPct')).toBeLessThan(firstProperties.indexOf('<a:buFont'));
    expect(firstProperties.indexOf('<a:buFont')).toBeLessThan(firstProperties.indexOf('<a:buAutoNum'));
    expect(firstProperties.indexOf('<a:buAutoNum')).toBeLessThan(firstProperties.indexOf('<a:tabLst>'));
    expect(firstProperties.indexOf('<a:tabLst>')).toBeLessThan(firstProperties.indexOf('<a:defRPr'));
    expect(firstProperties.indexOf('<a:defRPr')).toBeLessThan(firstProperties.indexOf('<a:extLst>'));
    expect(firstProperties).toContain('<a:buAutoNum type="alphaUcPeriod" startAt="2"/>');
    expect(firstProperties).toContain('<x:keep xmlns:x="urn:test">KEEP</x:keep>');
    expect(updated).toContain('<a:pPr marL="304800" custom="SECOND" algn="ctr"><a:buNone/></a:pPr>');
    expect(updated).toContain('<a:buChar char="▶"/>');
    expect(updated).not.toContain('Wingdings');
    expect(updated).not.toContain('hebrew2Minus');
    expect(updated).not.toContain('<a:buBlip>');
    expect(shape.richText.map(({ bullet }) => bullet)).toEqual([
      { kind: 'number', style: 'alphaUcPeriod', startAt: 2, indent: 20 },
      undefined,
      { kind: 'bullet', character: '▶', indent: 19 },
      undefined,
    ]);
  });

  it('reads and replaces list levels while keeping bullet hanging indents coherent', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        '<a:p><a:pPr lvl="2" marL="685800" indent="-228600" custom="FIRST"><a:spcBef><a:spcPts val="600"/></a:spcBef><a:buChar char="◆"/><a:tabLst><a:tab pos="500000" algn="l"/></a:tabLst><a:defRPr sz="1600"/><a:extLst><a:ext uri="urn:test"><x:keep xmlns:x="urn:test">KEEP</x:keep></a:ext></a:extLst></a:pPr><a:r><a:t>Level two</a:t></a:r></a:p><a:p><a:pPr lvl="0" marL="279400" indent="-279400" custom="SECOND"><a:buAutoNum type="romanUcPeriod" startAt="3"/></a:pPr><a:r><a:t>Root number</a:t></a:r></a:p><a:p><a:pPr lvl="1.5" marL="342900" indent="-342900" custom="THIRD"><a:buChar char="•"/><x:third xmlns:x="urn:test">THIRD</x:third></a:pPr><a:r><a:t>Malformed level</a:t></a:r></a:p><a:p><a:pPr lvl="1" marL="777778" indent="-111111" custom="FOURTH"><a:buChar char="►"/></a:pPr><a:r><a:t>Custom margin</a:t></a:r></a:p>',
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.richText.map(({ bullet, level }) => ({ bullet, level }))).toEqual([
      { bullet: { kind: 'bullet', character: '◆', indent: 18 }, level: 2 },
      {
        bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
        level: undefined,
      },
      { bullet: { kind: 'bullet', character: '•', indent: 27 }, level: undefined },
      { bullet: { kind: 'bullet', character: '►', indent: 30.62 }, level: 1 },
    ]);
    expect(pkg.mutations).toEqual(journal);

    shape.richText = [
      {
        runs: [{ text: 'Level three' }],
        bullet: { kind: 'bullet', character: '◆', indent: 18 },
        level: 3,
        spacing: { before: 6 },
        tabStops: [{ position: 500000 / 914400, alignment: 'left' }],
      },
      {
        runs: [{ text: 'Root number' }],
        bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
        level: 0,
      },
      { runs: [{ text: 'Cleared recognized bullet' }], bullet: false, level: 2, marginLeft: 0 },
      {
        runs: [{ text: 'Preserved custom margin' }],
        bullet: false,
        indent: -111111 / 12700,
        level: 0,
        marginLeft: 777778 / 12700,
      },
    ];

    const updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    const firstStart = updated.indexOf('<a:pPr lvl="3" marL="914400" indent="-228600" custom="FIRST">');
    const firstEnd = updated.indexOf('</a:pPr>', firstStart);
    const firstProperties = updated.slice(firstStart, firstEnd);
    expect(firstStart).toBeGreaterThan(-1);
    expect(firstProperties.indexOf('<a:spcBef>')).toBeLessThan(firstProperties.indexOf('<a:buSzPct'));
    expect(firstProperties.indexOf('<a:buChar')).toBeLessThan(firstProperties.indexOf('<a:tabLst>'));
    expect(firstProperties).toContain('<a:spcBef><a:spcPts val="600"/></a:spcBef>');
    expect(firstProperties).toContain('<x:keep xmlns:x="urn:test">KEEP</x:keep>');
    expect(updated).toContain('<a:pPr marL="279400" indent="-279400" custom="SECOND"><a:buSzPct');
    expect(updated).toContain('<a:pPr lvl="2" marL="0" custom="THIRD"><x:third xmlns:x="urn:test">THIRD</x:third><a:buNone/></a:pPr>');
    expect(updated).toContain('<a:pPr marL="777778" indent="-111111" custom="FOURTH"><a:buNone/></a:pPr>');
    expect(updated).not.toContain('lvl="1.5"');
    expect(shape.richText.map(({ bullet, level }) => ({ bullet, level }))).toEqual([
      { bullet: { kind: 'bullet', character: '◆', indent: 18 }, level: 3 },
      {
        bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
        level: undefined,
      },
      { bullet: undefined, level: 2 },
      { bullet: undefined, level: undefined },
    ]);
  });

  it('reads and replaces paragraph spacing without rebuilding unrelated paragraph properties', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        '<a:p><a:pPr algn="r" marL="228600" custom="FIRST"><a:lnSpc><a:spcPts val="2400"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef><a:spcAft><a:spcPts val="800"/></a:spcAft><a:buChar char="◆"/><a:tabLst><a:tab pos="500000" algn="l"/></a:tabLst><a:defRPr sz="1600"/><a:extLst><a:ext uri="urn:test"><x:keep xmlns:x="urn:test">KEEP</x:keep></a:ext></a:extLst></a:pPr><a:r><a:t>Exact</a:t></a:r></a:p><a:p><a:pPr custom="SECOND"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcBef><a:spcPct val="120000"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buNone/></a:pPr><a:r><a:t>Multiple</a:t></a:r></a:p><a:p><a:pPr custom="THIRD"><a:lnSpc><a:spcPts val="12x"/></a:lnSpc><a:spcBef><a:spcPts val="158401"/></a:spcBef><a:spcAft><a:spcPct val="100000"/></a:spcAft><x:third xmlns:x="urn:test">THIRD</x:third></a:pPr><a:r><a:t>Malformed</a:t></a:r></a:p><a:p><a:pPr custom="FOURTH"/><a:r><a:t>None</a:t></a:r></a:p>',
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.richText.map(({ spacing }) => spacing)).toEqual([
      { before: 6, after: 8, line: { kind: 'exact', points: 24 } },
      { line: { kind: 'multiple', factor: 1.5 } },
      undefined,
      undefined,
    ]);
    expect(pkg.mutations).toEqual(journal);

    shape.richText = [
      {
        runs: [{ text: 'Updated multiple' }],
        align: 'left',
        bullet: { kind: 'bullet', character: '◆', indent: 18 },
        spacing: { before: 7, after: 9, line: { kind: 'multiple', factor: 2 } },
        tabStops: [{ position: 500000 / 914400, alignment: 'left' }],
      },
      { runs: [{ text: 'Cleared' }], align: 'center', bullet: false, spacing: false },
      { runs: [{ text: 'Valid exact' }], bullet: false, spacing: { line: { kind: 'exact', points: 18 } } },
      { runs: [{ text: 'Still none' }], bullet: false },
    ];

    const updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    const firstStart = updated.indexOf('<a:pPr algn="l" marL="228600" custom="FIRST" indent="-228600">');
    const firstEnd = updated.indexOf('</a:pPr>', firstStart);
    const firstProperties = updated.slice(firstStart, firstEnd);
    expect(firstStart).toBeGreaterThan(-1);
    expect(firstProperties.indexOf('<a:lnSpc>')).toBeLessThan(firstProperties.indexOf('<a:spcBef>'));
    expect(firstProperties.indexOf('<a:spcBef>')).toBeLessThan(firstProperties.indexOf('<a:spcAft>'));
    expect(firstProperties.indexOf('<a:spcAft>')).toBeLessThan(firstProperties.indexOf('<a:buSzPct'));
    expect(firstProperties.indexOf('<a:buChar')).toBeLessThan(firstProperties.indexOf('<a:tabLst>'));
    expect(firstProperties.indexOf('<a:tabLst>')).toBeLessThan(firstProperties.indexOf('<a:defRPr'));
    expect(firstProperties.indexOf('<a:defRPr')).toBeLessThan(firstProperties.indexOf('<a:extLst>'));
    expect(firstProperties).toContain('<a:lnSpc><a:spcPct val="200000"/></a:lnSpc>');
    expect(firstProperties).toContain('<a:spcBef><a:spcPts val="700"/></a:spcBef>');
    expect(firstProperties).toContain('<a:spcAft><a:spcPts val="900"/></a:spcAft>');
    expect(firstProperties).toContain('<x:keep xmlns:x="urn:test">KEEP</x:keep>');
    expect(updated).toContain('<a:pPr custom="SECOND" algn="ctr"><a:buNone/></a:pPr>');
    expect(updated).toContain('<a:pPr custom="THIRD"><a:lnSpc><a:spcPts val="1800"/></a:lnSpc><x:third xmlns:x="urn:test">THIRD</x:third><a:buNone/></a:pPr>');
    expect(updated).not.toContain('12x');
    expect(updated).not.toContain('158401');
    expect(shape.richText.map(({ spacing }) => spacing)).toEqual([
      { before: 7, after: 9, line: { kind: 'multiple', factor: 2 } },
      undefined,
      { line: { kind: 'exact', points: 18 } },
      undefined,
    ]);
  });

  it('reads and replaces paragraph tab stops without rebuilding unrelated paragraph properties', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(
        '<a:p><a:r><a:t>First title</a:t></a:r></a:p>',
        '<a:p><a:pPr algn="r" lvl="1" marL="457200" indent="-228600" custom="FIRST"><a:spcBef><a:spcPts val="600"/></a:spcBef><a:buChar char="◆"/><a:tabLst><a:tab pos="1143000" algn="l"/><a:tab pos="-457200" algn="dec"/></a:tabLst><a:defRPr sz="1600"/><a:extLst><a:ext uri="urn:test"><x:keep xmlns:x="urn:test">KEEP</x:keep></a:ext></a:extLst></a:pPr><a:r><a:t>Valid tabs</a:t></a:r></a:p><a:p><a:pPr custom="SECOND"><a:tabLst/></a:pPr><a:r><a:t>Empty tabs</a:t></a:r></a:p><a:p><a:pPr custom="THIRD"><a:tabLst><a:tab pos="914400" algn="bad"/></a:tabLst><x:third xmlns:x="urn:test">THIRD</x:third></a:pPr><a:r><a:t>Malformed tabs</a:t></a:r></a:p><a:p><a:pPr custom="FOURTH"><a:tabLst><a:tab pos="914400" algn="l"/></a:tabLst><a:tabLst><a:tab pos="1828800" algn="r"/></a:tabLst></a:pPr><a:r><a:t>Duplicate lists</a:t></a:r></a:p>',
      ),
      part.contentType,
    );
    const shape = slide.shapes[0] as ShapeModel;
    const journal = [...pkg.mutations];

    expect(shape.richText.map(({ tabStops }) => tabStops)).toEqual([
      [
        { position: 1.25, alignment: 'left' },
        { position: -0.5, alignment: 'decimal' },
      ],
      [],
      undefined,
      undefined,
    ]);
    expect(pkg.mutations).toEqual(journal);

    shape.richText = [
      {
        runs: [{ text: 'Updated tabs' }],
        align: 'left',
        bullet: { kind: 'bullet', character: '◆', indent: 18 },
        level: 1,
        spacing: { before: 6 },
        tabStops: [
          { position: 2, alignment: 'right' },
          { position: 0, alignment: 'decimal' },
        ],
      },
      { runs: [{ text: 'Cleared empty' }], tabStops: false },
      { runs: [{ text: 'Explicit empty' }], tabStops: [] },
      { runs: [{ text: 'Cleared duplicates' }] },
    ];

    const updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    const firstStart = updated.indexOf('<a:pPr algn="l" lvl="1" marL="457200" indent="-228600" custom="FIRST">');
    const firstEnd = updated.indexOf('</a:pPr>', firstStart);
    const firstProperties = updated.slice(firstStart, firstEnd);
    expect(firstStart).toBeGreaterThan(-1);
    expect(firstProperties.indexOf('<a:spcBef>')).toBeLessThan(firstProperties.indexOf('<a:buSzPct'));
    expect(firstProperties.indexOf('<a:buChar')).toBeLessThan(firstProperties.indexOf('<a:tabLst>'));
    expect(firstProperties.indexOf('<a:tabLst>')).toBeLessThan(firstProperties.indexOf('<a:defRPr'));
    expect(firstProperties).toContain('<a:tab pos="1828800" algn="r"/><a:tab pos="0" algn="dec"/>');
    expect(firstProperties).toContain('<x:keep xmlns:x="urn:test">KEEP</x:keep>');
    expect(updated).toContain('<a:pPr custom="SECOND"><a:buNone/></a:pPr>');
    expect(updated).toContain('<a:pPr custom="THIRD"><a:buNone/><x:third xmlns:x="urn:test">THIRD</x:third><a:tabLst></a:tabLst></a:pPr>');
    expect(updated).toContain('<a:pPr custom="FOURTH"><a:buNone/></a:pPr>');
    expect(shape.richText.map(({ tabStops }) => tabStops)).toEqual([
      [
        { position: 2, alignment: 'right' },
        { position: 0, alignment: 'decimal' },
      ],
      undefined,
      [],
      undefined,
    ]);
  });

  it('does not mutate malformed text shapes that lack a text body or paragraph', async () => {
    for (const malformedTextBody of ['', '<p:txBody/>']) {
      const pkg = await OpcPackage.open(await modelFixture());
      const model = new PresentationModel(pkg);
      const slide = model.slides[1]!;
      const shape = slide.shapes[0] as ShapeModel;
      const part = pkg.requirePart(slide.partUri);
      pkg.setPart(
        part.uri,
        new TextDecoder().decode(part.bytes).replace(/<p:txBody>.*?<\/p:txBody>/, malformedTextBody),
        part.contentType,
      );
      const before = pkg.requirePart(part.uri).bytes;
      const journal = [...pkg.mutations];

      expect(() => {
        shape.text = 'Rejected';
      }).toThrow(malformedTextBody ? /text paragraph/ : /text body/);
      expect(() => {
        shape.richText = [{ runs: [{ text: 'Rejected' }] }];
      }).toThrow(malformedTextBody ? /text paragraph/ : /text body/);
      expect(() => shape.textMargins).toThrow(malformedTextBody ? /body properties/ : /text body/);
      expect(() => {
        shape.textMargins = 1;
      }).toThrow(malformedTextBody ? /body properties/ : /text body/);
      expect(() => shape.verticalAlignment).toThrow(malformedTextBody ? /body properties/ : /text body/);
      expect(() => {
        shape.verticalAlignment = 'bottom';
      }).toThrow(malformedTextBody ? /body properties/ : /text body/);
      expect(() => shape.textWrap).toThrow(malformedTextBody ? /body properties/ : /text body/);
      expect(() => {
        shape.textWrap = false;
      }).toThrow(malformedTextBody ? /body properties/ : /text body/);
      expect(() => shape.textDirection).toThrow(malformedTextBody ? /body properties/ : /text body/);
      expect(() => {
        shape.textDirection = 'vert';
      }).toThrow(malformedTextBody ? /body properties/ : /text body/);
      expect(() => shape.textFit).toThrow(malformedTextBody ? /body properties/ : /text body/);
      expect(() => {
        shape.textFit = 'shrink';
      }).toThrow(malformedTextBody ? /body properties/ : /text body/);
      expect(pkg.requirePart(part.uri).bytes).toEqual(before);
      expect(pkg.mutations).toEqual(journal);
    }
  });

  it('reads and edits slide size without changing notes, opaque XML, or shape transforms', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const shape = model.slides[1]!.shapes[0]!;
    const transform = shape.transform;
    const input = { width: inches(11.7), height: inches(8.3) };

    expect(model.slideSize).toEqual({ width: inches(10), height: inches(5.625) });
    expect(model.slideSize).not.toBe(model.slideSize);
    model.slideSize = input;
    input.width = inches(10);

    expect(model.slideSize).toEqual({ width: inches(11.7), height: inches(8.3) });
    expect(shape.transform).toEqual(transform);
    const updated = new TextDecoder().decode(pkg.requirePart('/ppt/presentation.xml').bytes);
    expect(updated).toContain('<p:sldSz cx="10698480" cy="7589520" type="screen" custom="KEEP"/>');
    expect(updated).toContain('<p:notesSz cx="5143500" cy="9144000"/>');
    expect(updated).toContain('<x:unknown xmlns:x="urn:test"/>');

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    expect(reopened.slideSize).toEqual({ width: inches(11.7), height: inches(8.3) });
  });

  it('inserts a missing slide size before notes and restores it after an outer rollback', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const part = pkg.requirePart('/ppt/presentation.xml');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(/<p:sldSz .*?\/>/, ''),
      part.contentType,
    );
    const model = new PresentationModel(pkg);
    expect(() => model.slideSize).toThrow(/slide size is missing/);

    model.slideSize = { width: inches(12), height: inches(7) };
    const inserted = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(inserted.indexOf('<p:sldSz')).toBeLessThan(inserted.indexOf('<p:notesSz'));
    expect(inserted).toContain('xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"');

    expect(() =>
      pkg.transaction(() => {
        model.slideSize = { width: inches(10), height: inches(10) };
        throw new Error('restore slide size');
      }),
    ).toThrow('restore slide size');
    expect(model.slideSize).toEqual({ width: inches(12), height: inches(7) });

    const insertedPart = pkg.requirePart(part.uri);
    pkg.setPart(
      insertedPart.uri,
      new TextDecoder().decode(insertedPart.bytes).replace(/<p:sldSz .*?\/>/, '<p:sldSz type="screen"/>'),
      insertedPart.contentType,
    );
    model.slideSize = { width: inches(10), height: inches(7.5) };
    expect(new TextDecoder().decode(pkg.requirePart(part.uri).bytes)).toContain(
      '<p:sldSz type="screen" cx="9144000" cy="6858000"/>',
    );
  });

  it('rejects invalid slide size edits without changing package bytes or mutations', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const before = pkg.requirePart('/ppt/presentation.xml').bytes;
    const journal = [...pkg.mutations];

    for (const value of [
      null,
      [],
      {},
      { width: Number.NaN, height: inches(1) },
      { width: inches(1), height: Number.POSITIVE_INFINITY },
      { width: inches(0.99), height: inches(1) },
      { width: inches(1), height: inches(56.01) },
    ]) {
      expect(() => {
        model.slideSize = value as never;
      }).toThrow();
    }
    expect(pkg.requirePart('/ppt/presentation.xml').bytes).toEqual(before);
    expect(pkg.mutations).toEqual(journal);

    const part = pkg.requirePart('/ppt/presentation.xml');
    pkg.setPart(
      part.uri,
      new TextDecoder().decode(part.bytes).replace(/<p:notesSz .*?\/>/, ''),
      part.contentType,
    );
    const withoutNotes = pkg.requirePart(part.uri).bytes;
    const withoutNotesJournal = [...pkg.mutations];
    expect(() => {
      model.slideSize = { width: inches(12), height: inches(7) };
    }).toThrow(/notes size is missing/);
    expect(pkg.requirePart(part.uri).bytes).toEqual(withoutNotes);
    expect(pkg.mutations).toEqual(withoutNotesJournal);

    const malformedPkg = await OpcPackage.open(await modelFixture());
    const malformedPart = malformedPkg.requirePart('/ppt/presentation.xml');
    malformedPkg.setPart(
      malformedPart.uri,
      new TextDecoder().decode(malformedPart.bytes).replace('cx="9144000"', 'cx="9144000.5"'),
      malformedPart.contentType,
    );
    expect(() => new PresentationModel(malformedPkg).slideSize).toThrow(/slide width is invalid/);
  });

  it('edits shape text and adds, duplicates, moves, and deletes slides with relationship updates', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const first = model.slides[1]!;
    (first.shapes[0] as ShapeModel).text = 'Edited first';
    expect(model.slides[1]!.title.text).toBe('Edited first');

    model.moveSlide(1, 0);
    expect(model.slides[0]).toBe(first);
    expect(model.slides.map(({ title }) => title.text)).toEqual(['Edited first', 'Second title']);
    const duplicate = model.duplicateSlide(0);
    expect(model.slides.find(({ partUri }) => partUri === duplicate.partUri)).toBe(duplicate);
    expect(duplicate.title.text).toBe('Edited first');
    expect((duplicate.shapes[1] as ImageModel).sourcePartUri).toBe('/ppt/media/image1.png');
    const blank = model.addSlide();
    expect(blank.shapes).toHaveLength(0);
    expect(model.slides).toHaveLength(4);

    model.deleteSlide(1);
    expect(model.slides).toHaveLength(3);
    expect(pkg.hasPart('/ppt/slides/slide2.xml')).toBe(false);
    expect(new TextDecoder().decode(pkg.requirePart('/ppt/presentation.xml').bytes)).toContain(
      '<x:unknown xmlns:x="urn:test"/>',
    );

    const reopened = new PresentationModel(await OpcPackage.open(await pkg.write()));
    expect(reopened.slides.map(({ title }) => title.text)).toEqual(['Edited first', 'Edited first', '']);
  });

  it('clones owned slide dependency subgraphs and garbage-collects only unreferenced owned parts', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const source = model.slides[1]!;
    const sourceImage = (source.shapes[1] as ImageModel).sourcePartUri!;
    const sourceChart = (source.shapes[3] as ChartModel).chartPartUri!;
    const sourceWorkbook = pkg.relationships(sourceChart)[0]!.resolvedTarget!;
    const sourceOpaque = pkg.relationships(source.partUri).find(({ id }) => id === 'rId3')!.resolvedTarget!;
    const sourceNotes = pkg.relationships(source.partUri).find(({ id }) => id === 'rId4')!.resolvedTarget!;
    const sourceNotesMaster = pkg
      .relationships(sourceNotes)
      .find(({ type }) => type.endsWith('/notesMaster'))!.resolvedTarget!;

    const duplicate = model.duplicateSlide(1);
    const duplicateImage = (duplicate.shapes[1] as ImageModel).sourcePartUri!;
    const duplicateChartModel = duplicate.shapes[3] as ChartModel;
    const duplicateChart = duplicateChartModel.chartPartUri!;
    const duplicateWorkbook = pkg.relationships(duplicateChart)[0]!.resolvedTarget!;
    const duplicateOpaque = pkg.relationships(duplicate.partUri).find(({ id }) => id === 'rId3')!.resolvedTarget!;
    const duplicateNotes = pkg.relationships(duplicate.partUri).find(({ id }) => id === 'rId4')!.resolvedTarget!;
    const duplicateNotesRelationships = pkg.relationships(duplicateNotes);
    const duplicateExternal = pkg.relationships(duplicate.partUri).find(({ id }) => id === 'rId5')!;

    expect(pkg.relationships(duplicate.partUri).map(({ id }) => id)).toEqual(
      pkg.relationships(source.partUri).map(({ id }) => id),
    );
    expect(duplicateImage).toBe(sourceImage);
    expect(duplicateChart).not.toBe(sourceChart);
    expect(duplicateWorkbook).not.toBe(sourceWorkbook);
    expect(pkg.requirePart(duplicateWorkbook).bytes).toEqual(pkg.requirePart(sourceWorkbook).bytes);
    expect(duplicateOpaque).toBe(sourceOpaque);
    expect(duplicateExternal).toMatchObject({ target: 'https://example.com', targetMode: 'External' });
    expect(duplicateNotes).not.toBe(sourceNotes);
    expect(duplicateNotesRelationships.find(({ type }) => type.endsWith('/slide'))?.resolvedTarget).toBe(
      duplicate.partUri,
    );
    expect(duplicateNotesRelationships.find(({ type }) => type.endsWith('/notesMaster'))?.resolvedTarget).toBe(
      sourceNotesMaster,
    );

    duplicateChartModel.setXml('<c:chartSpace xmlns:c="c"><c:chart><c:plotArea/></c:chart></c:chartSpace>');
    expect(new TextDecoder().decode(pkg.requirePart(sourceChart).bytes)).toContain('Sales');

    model.deleteSlide(model.slides.indexOf(duplicate));
    expect(pkg.hasPart(duplicateChart)).toBe(false);
    expect(pkg.hasPart(duplicateWorkbook)).toBe(false);
    expect(pkg.hasPart(duplicateNotes)).toBe(false);
    expect(pkg.hasPart(sourceChart)).toBe(true);
    expect(pkg.hasPart(sourceWorkbook)).toBe(true);
    expect(pkg.hasPart(sourceImage)).toBe(true);
    expect(pkg.hasPart(sourceOpaque)).toBe(true);
    expect(pkg.hasPart(sourceNotesMaster)).toBe(true);

    model.deleteSlide(model.slides.indexOf(source));
    expect(pkg.hasPart(sourceChart)).toBe(false);
    expect(pkg.hasPart(sourceWorkbook)).toBe(false);
    expect(pkg.hasPart(sourceNotes)).toBe(false);
    expect(pkg.hasPart(sourceImage)).toBe(true);
    expect(pkg.hasPart(sourceOpaque)).toBe(true);
    expect(pkg.hasPart(sourceNotesMaster)).toBe(true);

    const reopened = await OpcPackage.open(await pkg.write());
    expect(reopened.hasPart(sourceChart)).toBe(false);
    expect(new TextDecoder().decode(reopened.requirePart('/[Content_Types].xml').bytes)).not.toContain(
      '/ppt/charts/chart1.xml',
    );
  });

  it('clones shared image payloads on write while preserving shape identity', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const source = model.slides[1]!;
    const duplicate = model.duplicateSlide(1);
    const sourceImage = source.shapes[1] as ImageModel;
    const duplicateImage = duplicate.shapes[1] as ImageModel;
    const sharedPartUri = sourceImage.sourcePartUri!;
    const sourceBytes = pkg.requirePart(sharedPartUri).bytes;

    expect(duplicateImage.sourcePartUri).toBe(sharedPartUri);
    duplicateImage.replaceData(new Uint8Array([1, 2, 3]), 'image/png');
    expect(duplicate.shapes[1]).toBe(duplicateImage);
    expect(duplicateImage.sourcePartUri).not.toBe(sharedPartUri);
    expect(pkg.requirePart(sharedPartUri).bytes).toEqual(sourceBytes);
    expect(pkg.requirePart(duplicateImage.sourcePartUri!).bytes).toEqual(new Uint8Array([1, 2, 3]));
  });

  it('creates a new relationship when two image shapes share one rId', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const part = pkg.requirePart(slide.partUri);
    const xml = new TextDecoder().decode(part.bytes);
    const picture = xml.match(/<p:pic>.*?<\/p:pic>/)?.[0];
    expect(picture).toBeTruthy();
    pkg.setPart(
      slide.partUri,
      xml.replace('</p:spTree>', `${picture!.replace('id="3"', 'id="6"').replace('name="Image 1"', 'name="Image 2"')}</p:spTree>`),
      part.contentType,
    );
    const first = slide.shapes.find(({ id }) => id === 3) as ImageModel;
    const second = slide.shapes.find(({ id }) => id === 6) as ImageModel;
    const sharedPartUri = first.sourcePartUri!;

    second.replaceData(new Uint8Array([6, 6, 6]), 'image/png');
    expect(first.sourcePartUri).toBe(sharedPartUri);
    expect(second.sourcePartUri).toBeDefined();
    expect(second.sourcePartUri).not.toBe(sharedPartUri);
    expect(pkg.requirePart(second.sourcePartUri!).bytes).toEqual(new Uint8Array([6, 6, 6]));
    expect(pkg.relationships(slide.partUri).filter(({ resolvedTarget }) => resolvedTarget === sharedPartUri)).toHaveLength(1);
  });

  it('rolls back a shared image clone when package metadata cannot be updated', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const source = model.slides[1]!;
    const duplicate = model.duplicateSlide(1);
    const image = duplicate.shapes[1] as ImageModel;
    const sharedPartUri = (source.shapes[1] as ImageModel).sourcePartUri!;
    pkg.setPart('/[Content_Types].xml', '<invalid/>');
    const partUris = pkg.parts.map(({ uri }) => uri);
    const journal = [...pkg.mutations];

    expect(() => image.replaceData(new Uint8Array([4, 5, 6]), 'image/png')).toThrow(
      /Invalid \[Content_Types\]\.xml/,
    );
    expect(image.sourcePartUri).toBe(sharedPartUri);
    expect(pkg.parts.map(({ uri }) => uri)).toEqual(partUris);
    expect(pkg.mutations).toEqual(journal);
  });

  it('clones a shared chart and its owned workbook before raw XML editing', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const source = model.slides[1]!;
    const target = model.slides[0]!;
    const sourceChart = (source.shapes[3] as ChartModel).chartPartUri!;
    const sourceWorkbook = pkg.relationships(sourceChart)[0]!.resolvedTarget!;
    const targetPart = pkg.requirePart(target.partUri);
    const targetXml = new TextDecoder()
      .decode(targetPart.bytes)
      .replace(
        '<p:sld xmlns:p="p" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">',
        '<p:sld xmlns:p="p" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="r" xmlns:c="c">',
      )
      .replace(
        '</p:spTree>',
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Shared chart"/></p:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree>',
      );
    pkg.setPart(target.partUri, targetXml, targetPart.contentType);
    pkg.addRelationship(target.partUri, {
      id: 'rId1',
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
      target: '../charts/chart1.xml',
    });
    const chart = target.shapes.find(({ id }) => id === 3) as ChartModel;

    expect(chart.chartPartUri).toBe(sourceChart);
    chart.setXml('<c:chartSpace xmlns:c="c"><c:chart><c:plotArea/></c:chart></c:chartSpace>');
    const clonedChart = chart.chartPartUri!;
    expect(clonedChart).toBeDefined();
    const clonedWorkbook = pkg.relationships(clonedChart)[0]!.resolvedTarget!;
    expect(clonedWorkbook).toBeDefined();
    expect(target.shapes.find(({ id }) => id === 3)).toBe(chart);
    expect(clonedChart).not.toBe(sourceChart);
    expect(clonedWorkbook).not.toBe(sourceWorkbook);
    expect(pkg.requirePart(clonedWorkbook).bytes).toEqual(pkg.requirePart(sourceWorkbook).bytes);
    expect((source.shapes[3] as ChartModel).chartPartUri).toBe(sourceChart);
    expect(new TextDecoder().decode(pkg.requirePart(sourceChart).bytes)).toContain('Sales');
  });

  it('rolls back a partially cloned owned subgraph when a nested target is missing', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    pkg.setPart(
      '/ppt/charts/_rels/chart1.xml.rels',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/missing.xlsx"/></Relationships>',
    );
    const partUris = pkg.parts.map(({ uri }) => uri);
    const journal = [...pkg.mutations];

    expect(() => model.duplicateSlide(1)).toThrow(/Missing package part/);
    expect(pkg.parts.map(({ uri }) => uri)).toEqual(partUris);
    expect(pkg.mutations).toEqual(journal);
    expect(pkg.hasPart('/ppt/slides/slide3.xml')).toBe(false);
    expect(pkg.hasPart('/ppt/charts/chart2.xml')).toBe(false);
  });

  it('rolls back composite slide mutations when presentation XML fails after part creation', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    pkg.setPart('/ppt/presentation.xml', '<p:notPresentation xmlns:p="p"/>');
    const partUris = pkg.parts.map(({ uri }) => uri);
    const relationships = pkg.relationships('/ppt/presentation.xml');
    const journal = [...pkg.mutations];

    expect(() => model.addSlide()).toThrow(/Invalid presentation XML/);
    expect(pkg.parts.map(({ uri }) => uri)).toEqual(partUris);
    expect(pkg.relationships('/ppt/presentation.xml')).toEqual(relationships);
    expect(pkg.mutations).toEqual(journal);
    expect(pkg.hasPart('/ppt/slides/slide3.xml')).toBe(false);
  });

  it('keeps slide and shape identity through rollback while exposing live shape metadata', async () => {
    const pkg = await OpcPackage.open(await modelFixture());
    const model = new PresentationModel(pkg);
    const slide = model.slides[1]!;
    const shape = slide.shapes[0]!;
    const image = slide.shapes[1]!;

    expect(() =>
      pkg.transaction(() => {
        model.deleteSlide(1);
        expect(model.slides.includes(slide)).toBe(false);
        throw new Error('restore slide');
      }),
    ).toThrow('restore slide');
    expect(model.slides[1]).toBe(slide);
    expect(model.slides[1]!.shapes[0]).toBe(shape);

    const part = pkg.requirePart(slide.partUri);
    const renamed = new TextDecoder().decode(part.bytes).replace('name="Image 1"', 'name="Renamed image"');
    pkg.setPart(slide.partUri, renamed, part.contentType);
    expect(image.name).toBe('Renamed image');
    expect(slide.shapes[1]).toBe(image);

    const table = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Replacement table"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>Cell</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
    pkg.setPart(slide.partUri, renamed.replace(/<p:pic>.*?<\/p:pic>/, table), part.contentType);
    const replacement = slide.shapes.find(({ id }) => id === 3);
    expect(replacement).toBeInstanceOf(TableModel);
    expect(replacement).not.toBe(image);
  });
});
