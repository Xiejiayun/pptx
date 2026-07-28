import { readFile } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { PptxDocument, ShapeModel } from '@pptx/sdk';
import { importPptxGenJS } from './index.js';

interface PptxGenJSInstance {
  readonly version: string;
  layout: string;
  addSlide(): {
    addText(
      text: string | readonly { readonly text: string; readonly options?: Record<string, unknown> }[],
      options: Record<string, unknown>,
    ): void;
  };
  write(options: { outputType: 'uint8array'; compression: boolean }): Promise<Uint8Array>;
}

const require = createRequire(import.meta.url);
const PptxGenJS = require('pptxgenjs') as new () => PptxGenJSInstance;

describe('importPptxGenJS', () => {
  it('imports public PptxGenJS output and continues editing in the OOXML kernel', async () => {
    const generated = new PptxGenJS();
    expect(generated.version).toBe('4.0.1');
    generated.layout = 'LAYOUT_WIDE';
    const generatedSlide = generated.addSlide();
    generatedSlide.addText('Created by PptxGenJS', { x: 1, y: 1, w: 7, h: 1, align: 'center' });
    generatedSlide.addText(
      [
        {
          text: 'Bold red',
          options: { bold: true, fontFace: 'Aptos', fontSize: 24, color: 'ff0000' },
        },
        {
          text: 'italic',
          options: { italic: true, fontSize: 14, color: '4472C4', softBreakBefore: true },
        },
      ],
      { x: 1, y: 2, w: 7, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Left', options: { align: 'left' } },
        { text: 'Center', options: { align: 'center' } },
        { text: 'Right', options: { align: 'right' } },
        { text: 'Justify', options: { align: 'justify' } },
      ],
      { x: 1, y: 3, w: 7, h: 2, align: 'left' },
    );
    generatedSlide.addText('Standard\nSecond', { x: 1, y: 5, w: 3, h: 1, bullet: true });
    generatedSlide.addText('Custom', {
      x: 4,
      y: 5,
      w: 3,
      h: 1,
      bullet: { characterCode: '25BA', indent: 18 },
    });
    generatedSlide.addText('Public numberType', {
      x: 7,
      y: 5,
      w: 3,
      h: 1,
      bullet: { type: 'number', numberType: 'romanUcPeriod', numberStartAt: 3, indent: 22 },
    });
    generatedSlide.addText('Deprecated style', {
      x: 10,
      y: 5,
      w: 2,
      h: 1,
      bullet: { type: 'number', style: 'romanLcParenR', startAt: 4, indent: 24 },
    });
    generatedSlide.addText('Exact first\nExact second', {
      x: 1,
      y: 6,
      w: 3,
      h: 1,
      lineSpacing: 28,
      lineSpacingMultiple: 1.5,
      paraSpaceBefore: 6.25,
      paraSpaceAfter: 8.5,
    });
    generatedSlide.addText('Multiple', {
      x: 4,
      y: 6,
      w: 3,
      h: 1,
      lineSpacingMultiple: 1.5,
      paraSpaceBefore: 4.25,
      paraSpaceAfter: 7.75,
    });
    generatedSlide.addText('Zero spacing', {
      x: 7,
      y: 6,
      w: 3,
      h: 1,
      lineSpacing: 0,
      lineSpacingMultiple: 0,
      paraSpaceBefore: 0,
      paraSpaceAfter: 0,
    });
    generatedSlide.addText('Level one', {
      x: 10,
      y: 6,
      w: 2,
      h: 0.5,
      bullet: true,
      indentLevel: 1,
    });
    generatedSlide.addText('Custom level two', {
      x: 10,
      y: 6.5,
      w: 2,
      h: 0.5,
      bullet: { characterCode: '25BA', indent: 18 },
      indentLevel: 2,
    });
    generatedSlide.addText('Number level three', {
      x: 10,
      y: 7,
      w: 2,
      h: 0.5,
      bullet: { type: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
      indentLevel: 3,
    });
    generatedSlide.addText('No bullet level two', {
      x: 10,
      y: 7.5,
      w: 2,
      h: 0.5,
      indentLevel: 2,
    });
    generatedSlide.addText('Left\tCenter\tRight\tDecimal', {
      x: 1,
      y: 7.25,
      w: 8,
      h: 0.5,
      tabStops: [
        { position: 1 },
        { position: 2.25, alignment: 'ctr' },
        { position: 3.5, alignment: 'r' },
        { position: 4.75, alignment: 'dec' },
      ],
    });
    generatedSlide.addText('Empty tabs', {
      x: 1,
      y: 7.75,
      w: 3,
      h: 0.5,
      tabStops: [],
    });
    generatedSlide.addText(
      [
        { text: 'First\tA', options: { breakLine: true, tabStops: [{ position: 1.5, alignment: 'r' }] } },
        { text: 'Second\tB', options: { tabStops: [{ position: 2.5, alignment: 'ctr' }] } },
      ],
      { x: 5, y: 7.5, w: 4, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Single', options: { underline: true } },
        { text: ' Double', options: { underline: { style: 'dbl', color: 'ff0000' } } },
        { text: ' Wavy', options: { underline: { style: 'wavyDbl' } } },
        { text: ' None', options: { underline: { style: 'none' } } },
        { text: ' Dot dash', options: { underline: { style: 'dotDashHeavy' } } },
      ],
      { x: 9, y: 3, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'True', options: { strike: true } },
        { text: ' False', options: { strike: false } },
        { text: ' Single', options: { strike: 'sngStrike' } },
        { text: ' Double', options: { strike: 'dblStrike' } },
        { text: ' None', options: { strike: 'noStrike' } },
      ],
      { x: 9, y: 4, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Yellow', options: { highlight: 'ffff00' } },
        { text: ' Theme', options: { highlight: 'accent2' } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 5, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Red', options: { outline: { color: 'ff0000', size: 1.5 } } },
        { text: ' Theme', options: { outline: { color: 'accent1', size: 2 } } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 6, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Red', options: { glow: { color: 'ff0000', size: 8, opacity: 0.5 } } },
        { text: ' Theme', options: { glow: { color: 'accent1', size: 2.5, opacity: 1 } } },
        { text: ' Default', options: { glow: { size: 0, opacity: 0 } } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 7, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Super', options: { superscript: true } },
        { text: ' Sub', options: { subscript: true } },
        { text: ' Custom+', options: { baseline: 600 } },
        { text: ' Custom-', options: { baseline: -800 } },
        { text: ' Fraction', options: { baseline: 1.5 } },
        { text: ' Zero', options: { baseline: 0 } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 8, w: 3, h: 1 },
    );
    generatedSlide.addText(
      [
        { text: 'Positive', options: { charSpacing: 2.5 } },
        { text: ' Negative', options: { charSpacing: -1.25 } },
        { text: ' Fraction', options: { charSpacing: 0.004 } },
        { text: ' Zero', options: { charSpacing: 0 } },
        { text: ' Combined', options: { charSpacing: 3, baseline: 600 } },
        { text: ' None', options: {} },
      ],
      { x: 9, y: 9, w: 3, h: 1 },
    );
    generatedSlide.addText('Omitted margin', {
      x: 0,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Margin omitted',
    });
    generatedSlide.addText('Zero margin', {
      x: 0,
      y: 0.5,
      w: 2,
      h: 0.5,
      margin: 0,
      objectName: 'Margin zero',
    });
    generatedSlide.addText('Scalar margin', {
      x: 0,
      y: 1,
      w: 2,
      h: 0.5,
      margin: 10,
      objectName: 'Margin scalar',
    });
    generatedSlide.addText('Tuple margin', {
      x: 0,
      y: 1.5,
      w: 2,
      h: 0.5,
      margin: [4, 8, 8, 4],
      objectName: 'Margin tuple',
    });
    generatedSlide.addText('Fractional margin', {
      x: 0,
      y: 2,
      w: 2,
      h: 0.5,
      margin: 0.125,
      objectName: 'Margin fractional',
    });
    generatedSlide.addText('Negative margin', {
      x: 0,
      y: 2.5,
      w: 2,
      h: 0.5,
      margin: -0.5,
      objectName: 'Margin negative',
    });
    generatedSlide.addText('Asymmetric probe', {
      x: 0,
      y: 3,
      w: 2,
      h: 0.5,
      margin: [1, 2, 3, 4],
      objectName: 'Margin asymmetric probe',
    });
    generatedSlide.addText('Omitted vertical alignment', {
      x: 2,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Vertical omitted',
    });
    generatedSlide.addText('Top vertical alignment', {
      x: 2,
      y: 0.5,
      w: 2,
      h: 0.5,
      valign: 'top',
      objectName: 'Vertical top',
    });
    generatedSlide.addText('Middle vertical alignment', {
      x: 2,
      y: 1,
      w: 2,
      h: 0.5,
      valign: 'middle',
      objectName: 'Vertical middle',
    });
    generatedSlide.addText('Bottom vertical alignment', {
      x: 2,
      y: 1.5,
      w: 2,
      h: 0.5,
      valign: 'bottom',
      objectName: 'Vertical bottom',
    });
    generatedSlide.addText(
      [{ text: 'Ignored run vertical alignment', options: { valign: 'bottom' } }],
      {
        x: 2,
        y: 2,
        w: 2,
        h: 0.5,
        objectName: 'Vertical run ignored',
      },
    );
    generatedSlide.addText('Omitted text wrapping', {
      x: 4,
      y: 0,
      w: 2,
      h: 0.5,
      objectName: 'Wrap omitted',
    });
    generatedSlide.addText('Enabled text wrapping', {
      x: 4,
      y: 0.5,
      w: 2,
      h: 0.5,
      wrap: true,
      objectName: 'Wrap true',
    });
    generatedSlide.addText('Disabled text wrapping', {
      x: 4,
      y: 1,
      w: 2,
      h: 0.5,
      wrap: false,
      objectName: 'Wrap false',
    });
    generatedSlide.addText('Invalid text wrapping', {
      x: 4,
      y: 1.5,
      w: 2,
      h: 0.5,
      wrap: 'false',
      objectName: 'Wrap invalid fallback',
    });
    generatedSlide.addText(
      [{ text: 'Ignored run text wrapping', options: { wrap: false } }],
      {
        x: 4,
        y: 2,
        w: 2,
        h: 0.5,
        objectName: 'Wrap run ignored',
      },
    );
    const document = await importPptxGenJS(generated);
    expect(document.slides[0]?.title.text).toBe('Created by PptxGenJS');
    expect((document.slides[0]!.shapes[0] as ShapeModel).richText[0]!.align).toBe('center');
    const rich = document.slides[0]!.shapes[1] as ShapeModel;
    expect(rich.text).toBe('Bold red\nitalic');
    expect(rich.richText[0]!.runs).toEqual([
      {
        text: 'Bold red',
        style: {
          fontFamily: 'Aptos',
          fontSize: 24,
          bold: true,
          color: { kind: 'srgb', value: 'FF0000' },
        },
      },
      {
        text: 'italic',
        softBreakBefore: true,
        style: { fontSize: 14, italic: true, color: { kind: 'srgb', value: '4472C4' } },
      },
    ]);
    const aligned = document.slides[0]!.shapes[2] as ShapeModel;
    expect(aligned.richText.map(({ align }) => align)).toEqual(['left', 'center', 'right', 'justify']);
    aligned.richText = aligned.richText.map((paragraph, index) => ({
      runs: paragraph.runs,
      ...(index === 3
        ? { align: 'center' as const }
        : paragraph.align
          ? { align: paragraph.align }
          : {}),
    }));
    expect((document.slides[0]!.shapes[3] as ShapeModel).richText.map(({ bullet }) => bullet)).toEqual([
      { kind: 'bullet', character: '•', indent: 27 },
      { kind: 'bullet', character: '•', indent: 27 },
    ]);
    expect((document.slides[0]!.shapes[4] as ShapeModel).richText[0]!.bullet).toEqual({
      kind: 'bullet',
      character: '►',
      indent: 18,
    });
    expect((document.slides[0]!.shapes[5] as ShapeModel).richText[0]!.bullet).toEqual({
      kind: 'number',
      style: 'arabicPeriod',
      startAt: 3,
      indent: 22,
    });
    expect((document.slides[0]!.shapes[6] as ShapeModel).richText[0]!.bullet).toEqual({
      kind: 'number',
      style: 'romanLcParenR',
      startAt: 4,
      indent: 24,
    });
    expect((document.slides[0]!.shapes[7] as ShapeModel).richText.map(({ spacing }) => spacing)).toEqual([
      { before: 6.25, after: 8.5, line: { kind: 'exact', points: 28 } },
      { before: 6.25, after: 8.5, line: { kind: 'exact', points: 28 } },
    ]);
    expect((document.slides[0]!.shapes[8] as ShapeModel).richText[0]!.spacing).toEqual({
      before: 4.25,
      after: 7.75,
      line: { kind: 'multiple', factor: 1.5 },
    });
    expect((document.slides[0]!.shapes[9] as ShapeModel).richText[0]!.spacing).toBeUndefined();
    expect((document.slides[0]!.shapes[10] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'bullet', character: '•', indent: 27 },
      level: 1,
    });
    expect((document.slides[0]!.shapes[11] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'bullet', character: '►', indent: 18 },
      level: 2,
    });
    expect((document.slides[0]!.shapes[12] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'number', style: 'romanUcPeriod', startAt: 3, indent: 22 },
      level: 3,
    });
    const noBulletLevel = (document.slides[0]!.shapes[13] as ShapeModel).richText[0]!;
    expect(noBulletLevel.level).toBe(2);
    expect(noBulletLevel.bullet).toBeUndefined();
    expect((document.slides[0]!.shapes[14] as ShapeModel).richText[0]!.tabStops).toEqual([
      { position: 1, alignment: 'left' },
      { position: 2.25, alignment: 'center' },
      { position: 3.5, alignment: 'right' },
      { position: 4.75, alignment: 'decimal' },
    ]);
    expect((document.slides[0]!.shapes[15] as ShapeModel).richText[0]!.tabStops).toEqual([]);
    expect((document.slides[0]!.shapes[16] as ShapeModel).richText.map(({ tabStops }) => tabStops)).toEqual([
      [{ position: 1.5, alignment: 'right' }],
      [{ position: 2.5, alignment: 'center' }],
    ]);
    expect((document.slides[0]!.shapes[17] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.underline,
    )).toEqual([
      { style: 'sng' },
      { style: 'dbl', color: { kind: 'srgb', value: 'FF0000' } },
      { style: 'wavyDbl' },
      false,
      { style: 'dotDashHeavy' },
    ]);
    expect((document.slides[0]!.shapes[18] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.strike,
    )).toEqual(['sngStrike', undefined, 'sngStrike', 'dblStrike', false]);
    expect((document.slides[0]!.shapes[19] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.highlight,
    )).toEqual([
      { kind: 'srgb', value: 'FFFF00' },
      { kind: 'scheme', value: 'accent2' },
      undefined,
    ]);
    expect((document.slides[0]!.shapes[20] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.outline,
    )).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, size: 1.5 },
      { color: { kind: 'scheme', value: 'accent1' }, size: 2 },
      undefined,
    ]);
    expect((document.slides[0]!.shapes[21] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.glow,
    )).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, opacity: 0.5, size: 8 },
      { color: { kind: 'scheme', value: 'accent1' }, opacity: 1, size: 2.5 },
      { color: { kind: 'srgb', value: 'FFFFFF' }, opacity: 0, size: 0 },
      undefined,
    ]);
    expect((document.slides[0]!.shapes[22] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.baseline,
    )).toEqual(['superscript', 'subscript', 'superscript', 'subscript', 0.075, undefined, undefined]);
    const spaced = (document.slides[0]!.shapes[23] as ShapeModel).richText[0]!.runs;
    expect(spaced.map(({ style }) => style?.characterSpacing)).toEqual([2.5, -1.25, 0, undefined, 3, undefined]);
    expect(spaced[4]!.style!.baseline).toBe('superscript');
    const shapeByName = (name: string): ShapeModel => {
      const shape = document.slides[0]!.shapes.find((candidate) => candidate.name === name);
      expect(shape).toBeInstanceOf(ShapeModel);
      return shape as ShapeModel;
    };
    expect(shapeByName('Margin omitted').textMargins).toBeUndefined();
    expect(shapeByName('Margin zero').textMargins).toEqual({ top: 0, right: 0, bottom: 0, left: 0 });
    expect(shapeByName('Margin scalar').textMargins).toEqual({
      top: 10,
      right: 10,
      bottom: 10,
      left: 10,
    });
    expect(shapeByName('Margin tuple').textMargins).toEqual({ top: 4, right: 8, bottom: 8, left: 4 });
    expect(shapeByName('Margin fractional').textMargins).toEqual({
      top: 1_588 / 12_700,
      right: 1_588 / 12_700,
      bottom: 1_588 / 12_700,
      left: 1_588 / 12_700,
    });
    expect(shapeByName('Margin negative').textMargins).toEqual({
      top: -0.5,
      right: -0.5,
      bottom: -0.5,
      left: -0.5,
    });
    expect(shapeByName('Margin asymmetric probe').textMargins).toEqual({
      top: 4,
      right: 2,
      bottom: 3,
      left: 1,
    });
    expect([
      'Vertical omitted',
      'Vertical top',
      'Vertical middle',
      'Vertical bottom',
      'Vertical run ignored',
    ].map((name) => shapeByName(name).verticalAlignment)).toEqual([
      'middle',
      'top',
      'middle',
      'bottom',
      'middle',
    ]);
    expect([
      'Wrap omitted',
      'Wrap true',
      'Wrap false',
      'Wrap invalid fallback',
      'Wrap run ignored',
    ].map((name) => shapeByName(name).textWrap)).toEqual([
      true,
      true,
      false,
      true,
      true,
    ]);
    const importedXml = new TextDecoder().decode(
      document.opcPackage.requirePart(document.slides[0]!.partUri).bytes,
    );
    expect(importedXml).toContain('lIns="12700" tIns="50800" rIns="25400" bIns="38100"');
    expect(importedXml).toMatch(/name="Vertical omitted"[\s\S]*?<a:bodyPr[^>]*anchor="ctr"/);
    expect(importedXml).toMatch(/name="Vertical top"[\s\S]*?<a:bodyPr[^>]*anchor="t"/);
    expect(importedXml).toMatch(/name="Vertical middle"[\s\S]*?<a:bodyPr[^>]*anchor="ctr"/);
    expect(importedXml).toMatch(/name="Vertical bottom"[\s\S]*?<a:bodyPr[^>]*anchor="b"/);
    expect(importedXml).toMatch(/name="Vertical run ignored"[\s\S]*?<a:bodyPr[^>]*anchor="ctr"/);
    expect(importedXml).toMatch(/name="Wrap omitted"[\s\S]*?<a:bodyPr[^>]*wrap="square"/);
    expect(importedXml).toMatch(/name="Wrap true"[\s\S]*?<a:bodyPr[^>]*wrap="square"/);
    expect(importedXml).toMatch(/name="Wrap false"[\s\S]*?<a:bodyPr[^>]*wrap="none"/);
    expect(importedXml).toMatch(/name="Wrap invalid fallback"[\s\S]*?<a:bodyPr[^>]*wrap="square"/);
    expect(importedXml).toMatch(/name="Wrap run ignored"[\s\S]*?<a:bodyPr[^>]*wrap="square"/);
    document.slides[0]!.title.text = 'Edited by the OOXML kernel';
    document.duplicateSlide(0);

    const reopened = await PptxDocument.open(await document.write());
    expect(reopened.slides.map(({ title }) => title.text)).toEqual([
      'Edited by the OOXML kernel',
      'Edited by the OOXML kernel',
    ]);
    expect((reopened.slides[1]!.shapes[1] as ShapeModel).richText[0]!.runs[1]).toMatchObject({
      text: 'italic',
      softBreakBefore: true,
      style: { italic: true, color: { kind: 'srgb', value: '4472C4' } },
    });
    expect((reopened.slides[1]!.shapes[2] as ShapeModel).richText.map(({ align }) => align)).toEqual([
      'left',
      'center',
      'right',
      'center',
    ]);
    expect((reopened.slides[1]!.shapes[4] as ShapeModel).richText[0]!.bullet).toEqual({
      kind: 'bullet',
      character: '►',
      indent: 18,
    });
    expect((reopened.slides[1]!.shapes[6] as ShapeModel).richText[0]!.bullet).toMatchObject({
      kind: 'number',
      style: 'romanLcParenR',
      startAt: 4,
    });
    expect((reopened.slides[1]!.shapes[7] as ShapeModel).richText[0]!.spacing).toEqual({
      before: 6.25,
      after: 8.5,
      line: { kind: 'exact', points: 28 },
    });
    expect((reopened.slides[1]!.shapes[8] as ShapeModel).richText[0]!.spacing).toEqual({
      before: 4.25,
      after: 7.75,
      line: { kind: 'multiple', factor: 1.5 },
    });
    expect((reopened.slides[1]!.shapes[10] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'bullet', indent: 27 },
      level: 1,
    });
    expect((reopened.slides[1]!.shapes[12] as ShapeModel).richText[0]).toMatchObject({
      bullet: { kind: 'number', indent: 22 },
      level: 3,
    });
    expect((reopened.slides[1]!.shapes[14] as ShapeModel).richText[0]!.tabStops).toEqual([
      { position: 1, alignment: 'left' },
      { position: 2.25, alignment: 'center' },
      { position: 3.5, alignment: 'right' },
      { position: 4.75, alignment: 'decimal' },
    ]);
    expect((reopened.slides[1]!.shapes[16] as ShapeModel).richText.map(({ tabStops }) => tabStops)).toEqual([
      [{ position: 1.5, alignment: 'right' }],
      [{ position: 2.5, alignment: 'center' }],
    ]);
    expect((reopened.slides[1]!.shapes[17] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.underline,
    )).toEqual([
      { style: 'sng' },
      { style: 'dbl', color: { kind: 'srgb', value: 'FF0000' } },
      { style: 'wavyDbl' },
      false,
      { style: 'dotDashHeavy' },
    ]);
    expect((reopened.slides[1]!.shapes[18] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.strike,
    )).toEqual(['sngStrike', undefined, 'sngStrike', 'dblStrike', false]);
    expect((reopened.slides[1]!.shapes[19] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.highlight,
    )).toEqual([
      { kind: 'srgb', value: 'FFFF00' },
      { kind: 'scheme', value: 'accent2' },
      undefined,
    ]);
    expect((reopened.slides[1]!.shapes[20] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.outline,
    )).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, size: 1.5 },
      { color: { kind: 'scheme', value: 'accent1' }, size: 2 },
      undefined,
    ]);
    expect((reopened.slides[1]!.shapes[21] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.glow,
    )).toEqual([
      { color: { kind: 'srgb', value: 'FF0000' }, opacity: 0.5, size: 8 },
      { color: { kind: 'scheme', value: 'accent1' }, opacity: 1, size: 2.5 },
      { color: { kind: 'srgb', value: 'FFFFFF' }, opacity: 0, size: 0 },
      undefined,
    ]);
    expect((reopened.slides[1]!.shapes[22] as ShapeModel).richText[0]!.runs.map(
      ({ style }) => style?.baseline,
    )).toEqual(['superscript', 'subscript', 'superscript', 'subscript', 0.075, undefined, undefined]);
    const reopenedSpaced = (reopened.slides[1]!.shapes[23] as ShapeModel).richText[0]!.runs;
    expect(reopenedSpaced.map(({ style }) => style?.characterSpacing))
      .toEqual([2.5, -1.25, 0, undefined, 3, undefined]);
    expect(reopenedSpaced[4]!.style!.baseline).toBe('superscript');
    const reopenedMargins = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Margin '))
      .map(({ name, textMargins }) => [name, textMargins]);
    expect(reopenedMargins).toEqual([
      ['Margin omitted', undefined],
      ['Margin zero', { top: 0, right: 0, bottom: 0, left: 0 }],
      ['Margin scalar', { top: 10, right: 10, bottom: 10, left: 10 }],
      ['Margin tuple', { top: 4, right: 8, bottom: 8, left: 4 }],
      ['Margin fractional', {
        top: 1_588 / 12_700,
        right: 1_588 / 12_700,
        bottom: 1_588 / 12_700,
        left: 1_588 / 12_700,
      }],
      ['Margin negative', { top: -0.5, right: -0.5, bottom: -0.5, left: -0.5 }],
      ['Margin asymmetric probe', { top: 4, right: 2, bottom: 3, left: 1 }],
    ]);
    const reopenedVerticalAlignment = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Vertical '))
      .map(({ name, verticalAlignment }) => [name, verticalAlignment]);
    expect(reopenedVerticalAlignment).toEqual([
      ['Vertical omitted', 'middle'],
      ['Vertical top', 'top'],
      ['Vertical middle', 'middle'],
      ['Vertical bottom', 'bottom'],
      ['Vertical run ignored', 'middle'],
    ]);
    const reopenedWrapping = reopened.slides[1]!.shapes
      .filter((shape): shape is ShapeModel => shape instanceof ShapeModel)
      .filter(({ name }) => name.startsWith('Wrap '))
      .map(({ name, textWrap }) => [name, textWrap]);
    expect(reopenedWrapping).toEqual([
      ['Wrap omitted', true],
      ['Wrap true', true],
      ['Wrap false', false],
      ['Wrap invalid fallback', true],
      ['Wrap run ignored', true],
    ]);
  }, 10_000);

  it('keeps pptxgenjs out of every non-adapter package dependency list', async () => {
    const packagesDirectory = fileURLToPath(new URL('../..', import.meta.url));
    const packageNames = ['lossless-xml', 'model', 'opc', 'sdk', 'validator'];
    for (const packageName of packageNames) {
      const manifest = JSON.parse(await readFile(`${packagesDirectory}/${packageName}/package.json`, 'utf8')) as {
        dependencies?: Record<string, string>;
      };
      expect(manifest.dependencies?.pptxgenjs, packageName).toBeUndefined();
    }
  });
});
