import { createHash } from 'node:crypto';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Readable } from 'node:stream';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { openPptxStream, PptxDocument, ValidationError } from './index.js';

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
