import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import {
  ChartModel,
  ImageModel,
  PRESENTATION_FORMAT_PROFILES,
  PresentationModel,
  ShapeModel,
  TableModel,
  UnsupportedPresentationFormatError,
  emuToInches,
  inches,
  type PresentationFormat,
} from './index.js';

async function modelFixture(
  presentationContentType = PRESENTATION_FORMAT_PROFILES.pptx.presentationContentType,
): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="${presentationContentType}"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/ppt/embeddings/workbook1.xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/><Override PartName="/ppt/custom/opaque1.bin" ContentType="application/octet-stream"/><Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/><Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/></Types>`);
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="257" r:id="rId2"/><p:sldId id="256" r:id="rId1"/></p:sldIdLst><x:unknown xmlns:x="urn:test"/></p:presentation>');
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
    (shapes[1] as ImageModel).replaceData(new Uint8Array([1, 2, 3]), 'image/png');
    expect(model.opcPackage.requirePart('/ppt/media/image1.png').bytes).toEqual(new Uint8Array([1, 2, 3]));
    (shapes[3] as ChartModel).setXml('<c:chartSpace xmlns:c="c"><c:chart><c:plotArea/></c:chart></c:chartSpace>');
    expect((model.slides[1]!.shapes[3] as ChartModel).xml).toContain('<c:plotArea/>');
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
