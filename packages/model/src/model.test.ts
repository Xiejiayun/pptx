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
  zip.file('[Content_Types].xml', `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="${presentationContentType}"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>`);
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="257" r:id="rId2"/><p:sldId id="256" r:id="rId1"/></p:sldIdLst><x:unknown xmlns:x="urn:test"/></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r" xmlns:c="c"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm rot="60000"><a:off x="914400" y="1828800"/><a:ext cx="2743200" cy="914400"/></a:xfrm></p:spPr><p:txBody><a:p><a:r><a:t>First title</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id="3" name="Image 1"/></p:nvPicPr><p:blipFill><a:blip r:embed="rId1"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="20"/></a:xfrm></p:spPr></p:pic><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Table 1"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:p><a:r><a:t>A1</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>B1</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Chart 1"/></p:nvGraphicFramePr><a:graphic><a:graphicData><c:chart r:id="rId2"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>');
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>');
  zip.file('ppt/slides/slide2.xml', '<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 2"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Second title</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>');
  zip.file('ppt/media/image1.png', new Uint8Array([137, 80, 78, 71]));
  zip.file('ppt/charts/chart1.xml', '<c:chartSpace xmlns:c="c"><c:chart><c:plotArea><c:barChart><c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Sales</c:v></c:pt></c:strCache></c:strRef></c:tx><c:cat><c:strRef><c:strCache><c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>');
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
    expect(model.slides.map(({ title }) => title.text)).toEqual(['Second title', 'First title']);
    const shapes = model.slides[1]!.shapes;
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
    expect(emuToInches(model.slides[1]!.shapes[0]!.transform.x)).toBe(2);
    (shapes[2] as TableModel).setCellText(0, 1, 'Edited B1');
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
    expect(model.slides.map(({ title }) => title.text)).toEqual(['Edited first', 'Second title']);
    const duplicate = model.duplicateSlide(0);
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
});
