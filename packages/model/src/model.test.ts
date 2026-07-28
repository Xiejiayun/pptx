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
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="257" r:id="rId2"/><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500" type="screen" custom="KEEP"/><p:notesSz cx="5143500" cy="9144000"/><x:unknown xmlns:x="urn:test"/></p:presentation>');
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
    const imagePartUri = (shapes[1] as ImageModel).sourcePartUri;
    (shapes[1] as ImageModel).replaceData(new Uint8Array([1, 2, 3]), 'image/png');
    expect((shapes[1] as ImageModel).sourcePartUri).toBe(imagePartUri);
    expect(model.opcPackage.requirePart('/ppt/media/image1.png').bytes).toEqual(new Uint8Array([1, 2, 3]));
    const chartPartUri = (shapes[3] as ChartModel).chartPartUri;
    (shapes[3] as ChartModel).setXml('<c:chartSpace xmlns:c="c"><c:chart><c:plotArea/></c:chart></c:chartSpace>');
    expect((shapes[3] as ChartModel).chartPartUri).toBe(chartPartUri);
    expect((model.slides[1]!.shapes[3] as ChartModel).xml).toContain('<c:plotArea/>');
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

    expect(shape.richText.map(({ align }) => align)).toEqual(['left', 'right', undefined, undefined]);
    shape.richText = [
      { runs: [{ text: 'First' }], align: 'center' },
      { runs: [{ text: 'Second' }], align: 'justify' },
      { runs: [{ text: 'Third' }] },
      { runs: [{ text: 'Fourth' }], align: 'right' },
      { runs: [], align: 'left' },
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
      { runs: [{ text: 'No bullet' }], align: 'center' },
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
    expect(updated).toContain('<a:pPr marL="304800" indent="-100000" custom="SECOND" algn="ctr"><a:buNone/></a:pPr>');
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
      { runs: [{ text: 'Cleared recognized bullet' }], bullet: false, level: 2 },
      { runs: [{ text: 'Preserved custom margin' }], bullet: false, level: 0 },
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
    expect(updated).toContain('<a:pPr lvl="2" marL="0" indent="0" custom="THIRD"><x:third xmlns:x="urn:test">THIRD</x:third><a:buNone/></a:pPr>');
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
