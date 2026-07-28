import { createHash } from 'node:crypto';
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Readable } from 'node:stream';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
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
      { marginLeft: 0, runs: [] },
      {
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
      [{ runs: [{ text: 'x', breakLine: true }] }],
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
