import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument } from '@pptx/lossless-xml';
import { OpcPackage, relationshipPartUri } from '@pptx/opc';
import {
  ChartModel,
  ImageModel,
  ModelParseError,
  PRESENTATION_FORMAT_PROFILES,
  PRESET_SHAPE_TYPES,
  PresentationModel,
  ShapeModel,
  TableModel,
  UnsupportedPresentationFormatError,
  emuToInches,
  inches,
  type AddTableCellOptions,
  type AddTableCellInput,
  type AddTableOptions,
  type PresentationFormat,
  type ShapeArrows,
  type ShapeAdjustment,
  type ShapeArrowType,
  type ShapeFill,
  type ShapeLine,
  type ShapeLineDash,
  type ShapeShadow,
  type TableCellBorderInput,
  type TextBoxMarginInput,
} from './index.js';
import { readShapeHyperlink } from './shape-hyperlink.internal.js';
import { readShapeAdjustments } from './shape-adjustments.internal.js';
import { readSimpleShadow } from './simple-shadow.internal.js';

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

async function modelFixture(
  presentationContentType = PRESENTATION_FORMAT_PROFILES.pptx.presentationContentType,
): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="${presentationContentType}"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/ppt/embeddings/workbook1.xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/><Override PartName="/ppt/custom/opaque1.bin" ContentType="application/octet-stream"/><Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/><Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/></Types>`);
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="257" r:id="rId2"/><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500" type="screen" custom="KEEP"/><p:notesSz cx="5143500" cy="9144000"/><x:unknown xmlns:x="urn:test"/></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r" xmlns:c="c"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm rot="60000"><a:off x="914400" y="1828800"/><a:ext cx="2743200" cy="914400"/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:t>First title</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id="3" name="Image 1"/></p:nvPicPr><p:blipFill><a:blip r:embed="rId1"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="20"/></a:xfrm></p:spPr></p:pic><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>A1</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>B1</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Chart 1"/></p:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rId2"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>');
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/><Relationship Id="rId3" Type="urn:example:relationships/opaque" Target="../custom/opaque1.bin"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>');
  zip.file('ppt/slides/slide2.xml', '<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 2"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Second title</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>');
  zip.file('ppt/media/image1.png', new Uint8Array([137, 80, 78, 71]));
  zip.file('ppt/charts/chart1.xml', '<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:barChart><c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>');
  zip.file('ppt/charts/_rels/chart1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/workbook1.xlsx"/></Relationships>');
  zip.file('ppt/embeddings/workbook1.xlsx', new Uint8Array([80, 75, 3, 4, 1]));
  zip.file('ppt/custom/opaque1.bin', new Uint8Array([9, 8, 7]));
  zip.file('ppt/notesSlides/notesSlide1.xml', '<p:notes xmlns:p="p"><p:cSld><p:spTree/></p:cSld></p:notes>');
  zip.file('ppt/notesSlides/_rels/notesSlide1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/></Relationships>');
  zip.file('ppt/notesMasters/notesMaster1.xml', '<p:notesMaster xmlns:p="p"><p:cSld><p:spTree/></p:cSld></p:notesMaster>');
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

describe('PresentationModel', () => {
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
      [['line\nbreak']],
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

    table.setCellText(0, 3, 'Edited expanded');
    table.setCellTextDirection(0, 1, 'wordArtVert');
    table.setCellTextFit(0, 0, 'resize');
    table.setCellVerticalAlignment(0, 2, 'bottom');
    table.setTransform({ x: inches(1) });
    updated = new TextDecoder().decode(pkg.requirePart(part.uri).bytes);
    expect(updated).toContain('<a:t>Edited expanded</a:t>');
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
      .replace('<p:sld xmlns:p="p" xmlns:a="a">', '<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r" xmlns:c="c">')
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
