import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument } from '@pptx/lossless-xml';
import { OpcPackage, relativeRelationshipTarget } from '@pptx/opc';
import {
  CodecOwnershipError,
  CodecRegistry,
  GradientCodec,
  MasterLayoutThemeCodec,
  MediaCodec,
  type FeatureCodec,
  type GradientFill,
} from './index.js';

async function featureFixture(): Promise<OpcPackage> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:nvPr><p:ph type="title" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr><p:txBody><a:p><a:r><a:t>Slide title</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>');
  zip.file('ppt/slides/_rels/slide1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>');
  zip.file('ppt/slideLayouts/slideLayout1.xml', '<p:sldLayout xmlns:p="p" xmlns:a="a" name="Title"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Layout title"/><p:nvPr><p:ph type="title" idx="1"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Layout title</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sldLayout>');
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>');
  zip.file('ppt/slideMasters/slideMaster1.xml', '<p:sldMaster xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Master title"/><p:nvPr><p:ph type="title" idx="1"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Master title</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="1" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>');
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>');
  zip.file('ppt/theme/theme1.xml', '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"/></a:themeElements></a:theme>');
  return OpcPackage.open(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
}

describe('codec registry', () => {
  it('rejects equal-priority ownership conflicts', () => {
    const codec = (id: string): FeatureCodec<unknown> => ({
      id,
      ownership: { elements: ['x:item'] },
      detect: () => false,
      decode: () => undefined,
      encode: () => [],
      validate: () => [],
    });
    const registry = new CodecRegistry();
    registry.register(codec('first'));
    expect(() => registry.register(codec('second'))).toThrow(CodecOwnershipError);
  });
});

describe('GradientCodec', () => {
  const fill: GradientFill = {
    kind: 'linear-gradient',
    angle: 45,
    scaled: true,
    rotateWithShape: false,
    stops: [
      { offset: 0, color: '#2563EB' },
      { offset: 1, color: { source: { kind: 'scheme', value: 'accent2' }, alpha: 0.65, transforms: [{ kind: 'lumMod', value: 80000 }] } },
    ],
  };

  it('round-trips gradient stops, theme colors, transforms, and alpha', () => {
    const codec = new GradientCodec();
    const encoded = codec.encode(fill);
    const xml = LosslessXmlDocument.parse(encoded);
    const decoded = codec.decode(xml.elements('gradFill')[0]!, xml);
    expect(decoded.kind).toBe('linear-gradient');
    expect(decoded.stops[1]?.alpha).toBe(0.65);
    expect(decoded.stops[1]?.color).toMatchObject({ source: { kind: 'scheme', value: 'accent2' } });
  });

  it('sets slide backgrounds and shape fills with local XML patches', async () => {
    const pkg = await featureFixture();
    const codec = new GradientCodec();
    codec.setSlideBackground(pkg, '/ppt/slides/slide1.xml', fill);
    codec.setShapeFill(pkg, '/ppt/slides/slide1.xml', 2, fill);
    expect(codec.getSlideBackground(pkg, '/ppt/slides/slide1.xml')?.stops).toHaveLength(2);
    expect(codec.getShapeFill(pkg, '/ppt/slides/slide1.xml', 2)?.kind).toBe('linear-gradient');
  });
});

describe('MasterLayoutThemeCodec', () => {
  it('decodes, edits, copies, and relinks the inheritance chain', async () => {
    const pkg = await featureFixture();
    const codec = new MasterLayoutThemeCodec(pkg);
    const master = codec.masters[0]!;
    const layout = codec.layouts[0]!;
    const theme = codec.themes[0]!;
    expect(codec.masters[0]).toBe(master);
    expect(codec.layouts[0]).toBe(layout);
    expect(master.layouts[0]).toBe(layout);
    expect(codec.themes[0]).toBe(theme);
    expect(codec.presentationTheme).toBe(theme);
    expect(codec.presentationTheme).toBe(codec.presentationTheme);
    expect(master.theme).toBe(theme);
    expect(codec.masters).toHaveLength(1);
    expect(codec.layouts[0]?.placeholders[0]).toMatchObject({ type: 'title', index: 1 });
    expect(codec.themes[0]?.fonts).toEqual({ majorLatin: 'Aptos Display', minorLatin: 'Aptos' });
    theme.setFonts({ minorLatin: 'Noto Sans' });
    expect(theme.fonts).toEqual({ majorLatin: 'Aptos Display', minorLatin: 'Noto Sans' });
    codec.themes[0]!.setColor('accent1', '#2563EB');
    expect(codec.themes[0]?.colors.find(({ name }) => name === 'accent1')?.value).toBe('2563EB');
    expect(codec.materializeInheritedStyle('/ppt/slides/slide1.xml', 2)).toMatchObject({
      type: 'title',
      slide: { partUri: '/ppt/slides/slide1.xml' },
      layout: { partUri: '/ppt/slideLayouts/slideLayout1.xml' },
      master: { partUri: '/ppt/slideMasters/slideMaster1.xml' },
    });

    const copiedTheme = codec.copyTheme('/ppt/theme/theme1.xml');
    expect(codec.themes.find(({ partUri }) => partUri === copiedTheme.partUri)).toBe(copiedTheme);
    codec.relinkMasterTheme('/ppt/slideMasters/slideMaster1.xml', copiedTheme.partUri);
    pkg.updateRelationship('/ppt/presentation.xml', 'rId3', {
      target: relativeRelationshipTarget('/ppt/presentation.xml', copiedTheme.partUri),
    });
    expect(master.theme).toBe(copiedTheme);
    expect(codec.presentationTheme).toBe(copiedTheme);
    codec.deleteTheme('/ppt/theme/theme1.xml');
    expect(pkg.hasPart('/ppt/theme/theme1.xml')).toBe(false);
    const copiedLayout = codec.copyLayout('/ppt/slideLayouts/slideLayout1.xml');
    expect(codec.layouts.find(({ partUri }) => partUri === copiedLayout.partUri)).toBe(copiedLayout);
    expect(copiedLayout.masterPartUri).toBe(
      '/ppt/slideMasters/slideMaster1.xml',
    );
    codec.deleteLayout(copiedLayout.partUri);
    const copiedMaster = codec.copyMaster('/ppt/slideMasters/slideMaster1.xml');
    expect(codec.masters.find(({ partUri }) => partUri === copiedMaster.partUri)).toBe(copiedMaster);
    expect(copiedMaster.layouts.length).toBeGreaterThan(0);
    expect(codec.masters).toHaveLength(2);
    codec.deleteMaster(copiedMaster.partUri);
    expect(codec.masters).toHaveLength(1);

    const createdTheme = codec.createTheme(new TextDecoder().decode(pkg.requirePart(copiedTheme.partUri).bytes));
    const createdMaster = codec.createMaster(
      '<p:sldMaster xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree/></p:cSld></p:sldMaster>',
      createdTheme.partUri,
    );
    const createdLayout = codec.createLayout(
      createdMaster.partUri,
      '<p:sldLayout xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree/></p:cSld></p:sldLayout>',
    );
    expect(codec.masters.find(({ partUri }) => partUri === createdMaster.partUri)).toBe(createdMaster);
    expect(codec.themes.find(({ partUri }) => partUri === createdTheme.partUri)).toBe(createdTheme);
    expect(createdMaster.layouts.find(({ partUri }) => partUri === createdLayout.partUri)).toBe(createdLayout);
    expect(createdLayout.masterPartUri).toBe(createdMaster.partUri);
    codec.deleteMaster(createdMaster.partUri);
    codec.deleteTheme(createdTheme.partUri);
  });

  it('resolves only one safe presentation-direct theme relationship', async () => {
    const noDirectPackage = await featureFixture();
    noDirectPackage.removeRelationship('/ppt/presentation.xml', 'rId3');
    const noDirect = new MasterLayoutThemeCodec(noDirectPackage);
    expect(noDirect.masters[0]?.theme?.partUri).toBe('/ppt/theme/theme1.xml');
    expect(noDirect.presentationTheme).toBeUndefined();

    const duplicatePackage = await featureFixture();
    duplicatePackage.addRelationship('/ppt/presentation.xml', {
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
      target: 'theme/theme1.xml',
    });
    expect(new MasterLayoutThemeCodec(duplicatePackage).presentationTheme).toBeUndefined();

    const externalPackage = await featureFixture();
    externalPackage.updateRelationship('/ppt/presentation.xml', 'rId3', {
      target: 'https://example.com/theme.xml',
      targetMode: 'External',
    });
    expect(new MasterLayoutThemeCodec(externalPackage).presentationTheme).toBeUndefined();

    const danglingPackage = await featureFixture();
    const relationshipPart = danglingPackage.requirePart('/ppt/_rels/presentation.xml.rels');
    danglingPackage.setPart(
      relationshipPart.uri,
      new TextDecoder().decode(relationshipPart.bytes).replace(
        'Target="theme/theme1.xml"',
        'Target="theme/missing.xml"',
      ),
      relationshipPart.contentType,
    );
    expect(new MasterLayoutThemeCodec(danglingPackage).presentationTheme).toBeUndefined();

    const wrongTypePackage = await featureFixture();
    const wrongTypeBytes = wrongTypePackage.requirePart('/ppt/theme/theme1.xml').bytes;
    wrongTypePackage.setPart('/ppt/theme/theme1.xml', wrongTypeBytes, 'application/xml');
    expect(new MasterLayoutThemeCodec(wrongTypePackage).presentationTheme).toBeUndefined();

    const alternatePackage = await featureFixture();
    const alternateBytes = alternatePackage.requirePart('/ppt/theme/theme1.xml').bytes;
    alternatePackage.setPart(
      '/ppt/customThemes/primary.xml',
      alternateBytes,
      'application/vnd.openxmlformats-officedocument.theme+xml',
    );
    alternatePackage.setPart(
      '/ppt/theme/detached.xml',
      alternateBytes,
      'application/vnd.openxmlformats-officedocument.theme+xml',
    );
    alternatePackage.updateRelationship('/ppt/presentation.xml', 'rId3', {
      target: 'customThemes/primary.xml',
      targetMode: 'Internal',
    });
    const alternate = new MasterLayoutThemeCodec(alternatePackage);
    expect(alternate.themes).toHaveLength(3);
    expect(alternate.presentationTheme?.partUri).toBe('/ppt/customThemes/primary.xml');
    expect(alternate.presentationTheme?.fonts).toEqual({
      majorLatin: 'Aptos Display',
      minorLatin: 'Aptos',
    });
  });

  it('keeps theme font no-ops exact and rolls edits back with an outer transaction', async () => {
    const pkg = await featureFixture();
    const theme = new MasterLayoutThemeCodec(pkg).presentationTheme!;
    const before = pkg.requirePart(theme.partUri).bytes;
    const journal = [...pkg.mutations];

    theme.setFonts({ majorLatin: 'Aptos Display' });
    expect(pkg.requirePart(theme.partUri).bytes).toEqual(before);
    expect(pkg.mutations).toEqual(journal);

    expect(() =>
      pkg.transaction(() => {
        theme.setFonts({ majorLatin: 'Noto Sans Display', minorLatin: 'Noto Sans' });
        throw new Error('rollback theme fonts');
      }),
    ).toThrow('rollback theme fonts');
    expect(pkg.requirePart(theme.partUri).bytes).toEqual(before);
    expect(pkg.mutations).toEqual(journal);
    expect(theme.fonts).toEqual({ majorLatin: 'Aptos Display', minorLatin: 'Aptos' });
  });

  it('rolls back a master part when a dependent relationship cannot be created', async () => {
    const pkg = await featureFixture();
    const codec = new MasterLayoutThemeCodec(pkg);
    const partUris = pkg.parts.map(({ uri }) => uri);
    const journal = [...pkg.mutations];

    expect(() =>
      codec.createMaster(
        '<p:sldMaster xmlns:p="p"><p:cSld><p:spTree/></p:cSld></p:sldMaster>',
        '/ppt/theme/missing.xml',
      ),
    ).toThrow(/target part is missing/);
    expect(pkg.parts.map(({ uri }) => uri)).toEqual(partUris);
    expect(pkg.mutations).toEqual(journal);
  });
});

describe('MediaCodec', () => {
  it('embeds, deduplicates, lists, deletes, and diagnoses media', async () => {
    const pkg = await featureFixture();
    const codec = new MediaCodec(pkg);
    const bytes = new Uint8Array([0, 1, 2, 3, 4]);
    const first = await codec.addAudio('/ppt/slides/slide1.xml', bytes, {
      contentType: 'audio/mpeg',
      play: 'auto',
      loop: true,
      volume: 0.5,
    });
    const second = await codec.addAudio('/ppt/slides/slide1.xml', bytes, { contentType: 'audio/mpeg' });
    expect(second.mediaPartUri).toBe(first.mediaPartUri);
    const fromBlob = await codec.addAudio('/ppt/slides/slide1.xml', new Blob([bytes]), {
      contentType: 'audio/mpeg',
    });
    expect(fromBlob.mediaPartUri).toBe(first.mediaPartUri);
    const fromWebStream = await codec.addAudio(
      '/ppt/slides/slide1.xml',
      new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(bytes);
          controller.close();
        },
      }),
      { contentType: 'audio/mpeg' },
    );
    expect(fromWebStream.mediaPartUri).toBe(first.mediaPartUri);
    const external = await codec.addVideo('/ppt/slides/slide1.xml', 'https://example.com/video.mp4');
    const listed = codec.list('/ppt/slides/slide1.xml');
    expect(listed).toHaveLength(5);
    expect(listed[0]?.settings).toMatchObject({ play: 'auto', loop: true, volume: 0.5 });
    expect(codec.diagnostics(external, 'google-slides-import')[0]?.code).toBe('MEDIA_EXTERNAL_NOT_PORTABLE');
    codec.delete('/ppt/slides/slide1.xml', first.shapeId);
    expect(pkg.hasPart(first.mediaPartUri!)).toBe(true);
    codec.delete('/ppt/slides/slide1.xml', second.shapeId);
    codec.delete('/ppt/slides/slide1.xml', fromBlob.shapeId);
    codec.delete('/ppt/slides/slide1.xml', fromWebStream.shapeId);
    expect(pkg.hasPart(first.mediaPartUri!)).toBe(false);
  });

  it('rolls back media parts and relationships when the slide cannot accept a shape', async () => {
    const pkg = await featureFixture();
    pkg.setPart('/ppt/slides/slide1.xml', '<p:sld xmlns:p="p"/>');
    const codec = new MediaCodec(pkg);
    const partUris = pkg.parts.map(({ uri }) => uri);
    const relationships = pkg.relationships('/ppt/slides/slide1.xml');
    const journal = [...pkg.mutations];

    await expect(
      codec.addAudio('/ppt/slides/slide1.xml', new Uint8Array([1, 2, 3]), { contentType: 'audio/mpeg' }),
    ).rejects.toThrow(/no shape tree/);
    expect(pkg.parts.map(({ uri }) => uri)).toEqual(partUris);
    expect(pkg.relationships('/ppt/slides/slide1.xml')).toEqual(relationships);
    expect(pkg.mutations).toEqual(journal);
  });
});
