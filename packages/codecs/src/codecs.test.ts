import JSZip from 'jszip';
import { describe, expect, it, vi } from 'vitest';
import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
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
  it('creates canonical audio and video packages with stable names, metadata, and order', async () => {
    const pkg = await featureFixture();
    const originalSlide = pkg.requirePart('/ppt/slides/slide1.xml');
    pkg.setPart(
      originalSlide.uri,
      new TextDecoder().decode(originalSlide.bytes).replace(
        '</p:spTree>',
        '<p:extLst><p:ext uri="existing"/></p:extLst></p:spTree>',
      ),
      originalSlide.contentType,
    );
    const codec = new MediaCodec(pkg);
    const audioBytes = Uint8Array.of(0, 1, 2, 3, 4);
    const posterBytes = Uint8Array.of(5, 6, 7);
    const audio = await codec.addAudio('/ppt/slides/slide1.xml', audioBytes, {
      contentType: 'audio/mpeg',
      poster: posterBytes,
      posterContentType: 'image/png',
      play: 'auto',
      loop: true,
      volume: 0.5,
    });
    const video = await codec.addVideo('/ppt/slides/slide1.xml', Uint8Array.of(8, 9, 10), {
      contentType: 'video/mp4',
      fileName: 'clip.m4v',
      poster: Uint8Array.of(11, 12, 13),
      posterContentType: 'image/gif',
      x: -1,
      y: 0,
      width: 1,
      height: 2,
    });
    const named = await codec.addAudio('/ppt/slides/slide1.xml', Uint8Array.of(14), {
      name: 'Audio & narration',
      altText: 'Spoken "overview"',
      contentType: 'audio/wav',
      poster: Uint8Array.of(15),
      posterContentType: 'image/jpeg',
    });

    audioBytes[0] = 99;
    posterBytes[0] = 99;
    expect(pkg.requirePart(audio.mediaPartUri!)).toMatchObject({
      contentType: 'audio/mpeg',
      bytes: Uint8Array.of(0, 1, 2, 3, 4),
    });
    expect(pkg.requirePart(audio.posterPartUri!)).toMatchObject({
      contentType: 'image/png',
      bytes: Uint8Array.of(5, 6, 7),
    });
    expect(audio.mediaPartUri).toMatch(/\.mp3$/);
    expect(audio.posterPartUri).toMatch(/\.png$/);
    expect(video.mediaPartUri).toMatch(/\.m4v$/);
    expect(video.posterPartUri).toMatch(/\.gif$/);
    expect(pkg.requirePart(video.mediaPartUri!).contentType).toBe('video/mp4');
    expect(pkg.requirePart(named.mediaPartUri!).contentType).toBe('audio/wav');
    expect(named.mediaPartUri).toMatch(/\.wav$/);
    expect(named.posterPartUri).toMatch(/\.jpg$/);

    const xml = LosslessXmlDocument.parse(pkg.requirePart('/ppt/slides/slide1.xml').bytes);
    const shapeTree = xml.elements('spTree')[0]!;
    const directPictures = shapeTree.children.filter(
      (child): child is XmlElement => child.type === 'element' && child.localName === 'pic',
    );
    const directExtensionList = shapeTree.children.find(
      (child): child is XmlElement => child.type === 'element' && child.localName === 'extLst',
    );
    expect(directPictures).toHaveLength(3);
    expect(directPictures[2]!.end).toBe(directExtensionList!.start);
    expect(directPictures.map((picture) => {
      const properties = xml.descendants(picture).find(({ localName }) => localName === 'cNvPr')!;
      return [
        Number(xml.attribute(properties, 'id')?.value),
        xml.attribute(properties, 'name')?.value,
        xml.attribute(properties, 'descr')?.value,
      ];
    })).toEqual([
      [3, 'Media 0', undefined],
      [4, 'Media 1', undefined],
      [5, 'Audio & narration', 'Spoken "overview"'],
    ]);
    expect(xml.descendants(directPictures[0]!).map(({ localName }) => localName)).toEqual(
      expect.arrayContaining(['audioFile', 'media', 'picLocks', 'playback']),
    );
    expect(xml.descendants(directPictures[1]!).map(({ localName }) => localName)).toEqual(
      expect.arrayContaining(['videoFile', 'media', 'picLocks', 'playback']),
    );

    const relationships = pkg.relationships('/ppt/slides/slide1.xml');
    for (const model of [audio, video, named]) {
      const picture = directPictures.find((candidate) => {
        const properties = xml.descendants(candidate).find(({ localName }) => localName === 'cNvPr');
        return Number(properties ? xml.attribute(properties, 'id')?.value : -1) === model.shapeId;
      })!;
      const ids = xml.descendants(picture).flatMap(({ attributes }) => attributes)
        .filter(({ name }) => name === 'r:link' || name === 'r:embed')
        .map(({ value }) => value);
      expect(relationships.filter(({ id }) => ids.includes(id)).map(({ type }) => type)).toEqual([
        `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${model.kind}`,
        'http://schemas.microsoft.com/office/2007/relationships/media',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      ]);
    }

    const listed = [...codec.list('/ppt/slides/slide1.xml')]
      .sort((left, right) => left.shapeId - right.shapeId);
    expect(listed).toHaveLength(3);
    expect(listed).toMatchObject([
      {
        kind: 'audio',
        shapeId: 3,
        mediaPartUri: audio.mediaPartUri,
        posterPartUri: audio.posterPartUri,
        settings: { play: 'auto', loop: true, hideWhenStopped: false, volume: 0.5 },
      },
      {
        kind: 'video',
        shapeId: 4,
        mediaPartUri: video.mediaPartUri,
        posterPartUri: video.posterPartUri,
      },
      {
        kind: 'audio',
        shapeId: 5,
        mediaPartUri: named.mediaPartUri,
        posterPartUri: named.posterPartUri,
      },
    ]);
  });

  it('deduplicates by bytes and exact content type and collects only final references', async () => {
    const pkg = await featureFixture();
    const codec = new MediaCodec(pkg);
    const mediaBytes = Uint8Array.of(1, 2, 3);
    const posterBytes = Uint8Array.of(4, 5, 6);
    const first = await codec.addVideo('/ppt/slides/slide1.xml', mediaBytes, {
      contentType: 'video/mp4',
      fileName: 'first.mp4',
      poster: posterBytes,
      posterContentType: 'image/png',
    });
    const shared = await codec.addVideo('/ppt/slides/slide1.xml', mediaBytes, {
      contentType: 'video/mp4',
      fileName: 'second.m4v',
      poster: posterBytes,
      posterContentType: 'image/png',
    });
    const differentType = await codec.addAudio('/ppt/slides/slide1.xml', mediaBytes, {
      contentType: 'audio/ogg',
      poster: posterBytes,
      posterContentType: 'image/gif',
    });
    expect(shared.mediaPartUri).toBe(first.mediaPartUri);
    expect(shared.posterPartUri).toBe(first.posterPartUri);
    expect(differentType.mediaPartUri).not.toBe(first.mediaPartUri);
    expect(differentType.posterPartUri).not.toBe(first.posterPartUri);

    codec.delete('/ppt/slides/slide1.xml', first.shapeId);
    expect(pkg.hasPart(first.mediaPartUri!)).toBe(true);
    expect(pkg.hasPart(first.posterPartUri!)).toBe(true);
    codec.delete('/ppt/slides/slide1.xml', shared.shapeId);
    expect(pkg.hasPart(first.mediaPartUri!)).toBe(false);
    expect(pkg.hasPart(first.posterPartUri!)).toBe(false);
    expect(pkg.hasPart(differentType.mediaPartUri!)).toBe(true);
    expect(pkg.hasPart(differentType.posterPartUri!)).toBe(true);

    const external = await codec.addVideo(
      '/ppt/slides/slide1.xml',
      'https://example.com/video.mp4',
      { poster: Uint8Array.of(7, 8, 9), posterContentType: 'image/jpeg' },
    );
    const relationshipCount = pkg.relationships('/ppt/slides/slide1.xml').length;
    expect(codec.diagnostics(external, 'google-slides-import')[0]?.code)
      .toBe('MEDIA_EXTERNAL_NOT_PORTABLE');
    codec.delete('/ppt/slides/slide1.xml', external.shapeId);
    expect(pkg.relationships('/ppt/slides/slide1.xml')).toHaveLength(relationshipCount - 2);
    expect(pkg.hasPart(external.posterPartUri!)).toBe(false);
    expect(pkg.hasPart(differentType.mediaPartUri!)).toBe(true);
    codec.delete('/ppt/slides/slide1.xml', differentType.shapeId);
    expect(pkg.hasPart(differentType.mediaPartUri!)).toBe(false);
    expect(pkg.hasPart(differentType.posterPartUri!)).toBe(false);
  });

  it('leaves the complete package unchanged after input, source, transcode, and poster failures', async () => {
    const pkg = await featureFixture();
    const codec = new MediaCodec(pkg);
    const before = await packageSnapshot(pkg);
    const failingStream = (): AsyncIterable<Uint8Array> => ({
      async *[Symbol.asyncIterator]() {
        yield Uint8Array.of(1);
        throw new Error('stream failed');
      },
    });
    const failures: Array<() => Promise<unknown>> = [
      () => codec.addAudio('/ppt/slides/slide1.xml', Uint8Array.of(1), { volume: 2 }),
      () => codec.addAudio('/ppt/slides/slide1.xml', 'data:audio/mpeg;base64,A===', {}),
      () => codec.addAudio('/ppt/slides/slide1.xml', new Uint8Array(), {}),
      () => codec.addAudio('/ppt/slides/slide1.xml', Uint8Array.of(1), {
        contentType: 'audio/mpeg',
        fileName: 'voice.wav',
      }),
      () => codec.addAudio('/ppt/slides/slide1.xml', '/__pptx_missing__/voice.mp3', {}),
      () => codec.addAudio('/ppt/slides/slide1.xml', failingStream(), {}),
      () => codec.addAudio('/ppt/slides/slide1.xml', Uint8Array.of(1), {
        transcode: async () => { throw new Error('transcode failed'); },
      }),
      () => codec.addVideo('/ppt/slides/slide1.xml', Uint8Array.of(1), {
        poster: 'data:image/png;base64,A===',
      }),
      () => codec.addVideo('/ppt/slides/slide1.xml', Uint8Array.of(1), {
        poster: failingStream(),
      }),
    ];

    for (const fail of failures) {
      await expect(fail()).rejects.toThrow();
      expect(await packageSnapshot(pkg)).toEqual(before);
    }
  });

  it('rolls back shape, allocation, XML, relationship, and outer-transaction failures', async () => {
    const missingTree = await featureFixture();
    missingTree.setPart('/ppt/slides/slide1.xml', '<p:sld xmlns:p="p"/>');
    const missingTreeBefore = await packageSnapshot(missingTree);
    await expect(new MediaCodec(missingTree).addAudio(
      '/ppt/slides/slide1.xml',
      Uint8Array.of(1, 2, 3),
      { contentType: 'audio/mpeg' },
    )).rejects.toThrow(/shape tree/i);
    expect(await packageSnapshot(missingTree)).toEqual(missingTreeBefore);

    const allocation = await featureFixture();
    const allocationBefore = await packageSnapshot(allocation);
    const allocate = vi.spyOn(allocation, 'allocatePartUri')
      .mockImplementationOnce(() => '/ppt/media/media1.mp3')
      .mockImplementationOnce(() => { throw new Error('allocation failed'); });
    await expect(new MediaCodec(allocation).addAudio(
      '/ppt/slides/slide1.xml',
      Uint8Array.of(1, 2, 3),
      { contentType: 'audio/mpeg' },
    )).rejects.toThrow('allocation failed');
    allocate.mockRestore();
    expect(await packageSnapshot(allocation)).toEqual(allocationBefore);

    const invalidXml = await featureFixture();
    const invalidXmlBefore = await packageSnapshot(invalidXml);
    const originalAppend = LosslessXmlDocument.prototype.appendChildXml;
    const append = vi.spyOn(LosslessXmlDocument.prototype, 'appendChildXml')
      .mockImplementation(function (this: LosslessXmlDocument, element, value) {
        if (element.localName === 'spTree') throw new Error('invalid XML append');
        return originalAppend.call(this, element, value);
      });
    await expect(new MediaCodec(invalidXml).addAudio(
      '/ppt/slides/slide1.xml',
      Uint8Array.of(1, 2, 3),
      { contentType: 'audio/mpeg' },
    )).rejects.toThrow('invalid XML append');
    append.mockRestore();
    expect(await packageSnapshot(invalidXml)).toEqual(invalidXmlBefore);

    const relationship = await featureFixture();
    const relationshipBefore = await packageSnapshot(relationship);
    const originalAddRelationship = relationship.addRelationship.bind(relationship);
    let relationshipCalls = 0;
    const addRelationship = vi.spyOn(relationship, 'addRelationship')
      .mockImplementation((sourcePartUri, input) => {
        relationshipCalls += 1;
        if (relationshipCalls === 2) throw new Error('relationship target failed');
        return originalAddRelationship(sourcePartUri, input);
      });
    await expect(new MediaCodec(relationship).addAudio(
      '/ppt/slides/slide1.xml',
      Uint8Array.of(1, 2, 3),
      { contentType: 'audio/mpeg' },
    )).rejects.toThrow('relationship target failed');
    addRelationship.mockRestore();
    expect(await packageSnapshot(relationship)).toEqual(relationshipBefore);

    const outer = await featureFixture();
    const outerCodec = new MediaCodec(outer);
    const created = await outerCodec.addAudio(
      '/ppt/slides/slide1.xml',
      Uint8Array.of(1, 2, 3),
      { contentType: 'audio/mpeg' },
    );
    const outerBefore = await packageSnapshot(outer);
    expect(() => outer.transaction(() => {
      outerCodec.delete('/ppt/slides/slide1.xml', created.shapeId);
      throw new Error('outer rollback');
    })).toThrow('outer rollback');
    expect(await packageSnapshot(outer)).toEqual(outerBefore);

    const deletion = await featureFixture();
    const deletionCodec = new MediaCodec(deletion);
    const removable = await deletionCodec.addAudio(
      '/ppt/slides/slide1.xml',
      Uint8Array.of(4, 5, 6),
      { contentType: 'audio/mpeg' },
    );
    const deletionBefore = await packageSnapshot(deletion);
    const originalRemoveRelationship = deletion.removeRelationship.bind(deletion);
    let removalCalls = 0;
    const removeRelationship = vi.spyOn(deletion, 'removeRelationship')
      .mockImplementation((sourcePartUri, id) => {
        removalCalls += 1;
        if (removalCalls === 2) throw new Error('relationship removal failed');
        return originalRemoveRelationship(sourcePartUri, id);
      });
    expect(() => deletionCodec.delete('/ppt/slides/slide1.xml', removable.shapeId))
      .toThrow('relationship removal failed');
    removeRelationship.mockRestore();
    expect(await packageSnapshot(deletion)).toEqual(deletionBefore);
  });

  it('rolls back media source replacement after relationship and payload mutation', async () => {
    const pkg = await featureFixture();
    const codec = new MediaCodec(pkg);
    const media = await codec.addAudio(
      '/ppt/slides/slide1.xml',
      Uint8Array.of(1, 2, 3),
      { contentType: 'audio/mpeg' },
    );
    const before = await packageSnapshot(pkg);
    const originalSetPart = pkg.setPart.bind(pkg);
    const setPart = vi.spyOn(pkg, 'setPart').mockImplementation((uri, bytes, contentType) => {
      if (uri === '/ppt/slides/slide1.xml') throw new Error('replacement slide write failed');
      return originalSetPart(uri, bytes, contentType);
    });
    await expect(codec.replaceSource(
      '/ppt/slides/slide1.xml',
      media.shapeId,
      'audio',
      Uint8Array.of(4, 5, 6),
      { contentType: 'audio/mpeg' },
    )).rejects.toThrow('replacement slide write failed');
    setPart.mockRestore();
    expect(await packageSnapshot(pkg)).toEqual(before);
    await expect(codec.replaceSource(
      '/ppt/slides/slide1.xml',
      media.shapeId,
      'video',
      Uint8Array.of(7, 8),
      { contentType: 'video/mp4' },
    )).rejects.toThrow(/is not video/);
    expect(await packageSnapshot(pkg)).toEqual(before);
    const posterSetPart = vi.spyOn(pkg, 'setPart').mockImplementation((uri, bytes, contentType) => {
      if (uri === '/ppt/slides/slide1.xml') throw new Error('poster slide write failed');
      return originalSetPart(uri, bytes, contentType);
    });
    await expect(codec.replacePoster(
      '/ppt/slides/slide1.xml',
      media.shapeId,
      Uint8Array.of(9, 10),
      { contentType: 'image/gif' },
    )).rejects.toThrow('poster slide write failed');
    posterSetPart.mockRestore();
    expect(await packageSnapshot(pkg)).toEqual(before);
  });
});

async function packageSnapshot(pkg: OpcPackage): Promise<unknown> {
  const output = await OpcPackage.open(await pkg.write());
  return {
    ...packageState(pkg),
    output: packageState(output),
    journal: pkg.mutations.map((mutation) => ({ ...mutation })),
  };
}

function packageState(pkg: OpcPackage) {
  const partSources = pkg.parts
    .filter(({ uri }) => !uri.endsWith('.rels'))
    .map(({ uri }) => uri);
  return {
    parts: pkg.parts.map(({ uri, contentType, bytes }) => ({
      uri,
      contentType,
      bytes: new Uint8Array(bytes),
    })),
    relationships: ['/', ...partSources].map((uri) => [
      uri,
      pkg.relationships(uri).map((relationship) => ({ ...relationship })),
    ]),
    graph: pkg.graph,
    slide: new Uint8Array(pkg.requirePart('/ppt/slides/slide1.xml').bytes),
  };
}
