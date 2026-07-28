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
        runs: [
          {
            text: 'Bold & ',
            style: {
              fontFamily: 'Aptos & Display',
              fontSize: 12.5,
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
              italic: true,
              color: { kind: 'scheme', value: 'accent1' },
            },
          },
        ],
      },
      { runs: [] },
      {
        runs: [
          {
            text: 'Last',
            style: { fontFamily: '+mn-lt', color: { kind: 'scheme', value: 'tx1' } },
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
