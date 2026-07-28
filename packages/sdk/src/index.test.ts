import { createHash } from 'node:crypto';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Readable } from 'node:stream';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { PRESENTATION_FORMAT_PROFILES, type PresentationFormat } from '@pptx/model';
import { OpcPackage } from '@pptx/opc';
import { validatePackage } from '@pptx/validator';
import {
  degrees,
  inches,
  ModelParseError,
  openPptxStream,
  PptxDocument,
  ShapeModel,
  ValidationError,
} from './index.js';

async function titleFixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="p" xmlns:r="r"><p:sldIdLst><p:sldId id="256" r:id="rId7"/></p:sldIdLst></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Original</a:t></a:r></a:p></p:txBody><x:unknown xmlns:x="x" custom="keep"/></p:sp></p:spTree></p:cSld></p:sld>');
  zip.file('ppt/theme/theme1.xml', '<a:theme xmlns:a="a"><x:opaque xmlns:x="x">KEEP</x:opaque></a:theme>');
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

describe('PptxDocument vertical slice', () => {
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
    const presentationXml = new TextDecoder().decode(
      document.opcPackage.requirePart('/ppt/presentation.xml').bytes,
    );
    expect(presentationXml).toContain('<p:sldSz cx="10698480" cy="7589520"/>');
    expect(presentationXml).toContain('<p:notesSz cx="7589520" cy="10698480"/>');

    document.addSlide();
    const reopened = await PptxDocument.open(await document.write());
    expect(validatePackage(reopened.opcPackage).filter(({ severity }) => severity === 'error')).toEqual([]);
    expect(new TextDecoder().decode(reopened.opcPackage.requirePart('/ppt/presentation.xml').bytes)).toContain(
      '<p:sldSz cx="10698480" cy="7589520"/>',
    );
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
    expect(() => slide.addText('two\nlines')).toThrow(/do not support line breaks/);
    expect(() => {
      first.text = 'two\nlines';
    }).toThrow(/do not support line breaks/);
    expect(() => slide.addText('invalid\u0000xml')).toThrow(/invalid XML characters/);
    expect(() => slide.addText('bad width', { width: 0 as never })).toThrow(/width must be greater/);
    expect(() => slide.addText('bad coordinate', { x: Number.NaN as never })).toThrow(/x must be finite/);
    expect(() => slide.addText('bad flip', { flipHorizontal: 'yes' as never })).toThrow(/must be a boolean/);
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
    expect(slideXml).toContain('<a:t>Updated &amp; preserved</a:t>');
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
