import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import {
  normalizeSlideNumberOptions,
  readSlideNumber,
  replaceSlideNumber,
  replaceSlideNumberCachedText,
} from './slide-number.internal.js';
import type { SlideNumberOptions } from './slide-number.js';

const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';

describe('slide number values', () => {
  it('normalizes defaults, explicit zero, quantization, colors, margins, and style', () => {
    expect(normalizeSlideNumberOptions({})).toEqual({
      x: 0,
      y: 0,
      width: 800_000,
      height: 300_000,
      align: 'left',
      rtl: false,
      style: { lang: 'en-US', bold: false, italic: false },
    });

    expect(normalizeSlideNumberOptions({
      x: 0.49,
      y: -10.51,
      width: 1.49,
      height: 2.51,
      align: 'justify',
      rtl: true,
      valign: 'middle',
      margin: [1, 2, 3, 4],
      style: {
        fontFamily: 'Aptos',
        fontSize: 18.126,
        lang: 'zh-CN',
        bold: true,
        italic: true,
        color: { kind: 'srgb', value: 'ff3399' },
        transparency: 12.3456,
      },
    })).toEqual({
      x: 0,
      y: -11,
      width: 1,
      height: 3,
      align: 'justify',
      rtl: true,
      valign: 'middle',
      margin: { top: 1, right: 2, bottom: 3, left: 4 },
      style: {
        fontFamily: 'Aptos',
        fontSize: 18.13,
        lang: 'zh-CN',
        bold: true,
        italic: true,
        color: { kind: 'srgb', value: 'FF3399' },
        transparency: 12.346,
      },
    });

    expect(normalizeSlideNumberOptions({
      margin: 0,
      style: { transparency: 25 },
    })).toMatchObject({
      margin: { top: 0, right: 0, bottom: 0, left: 0 },
      style: {
        color: { kind: 'scheme', value: 'tx1' },
        transparency: 25,
      },
    });
    expect(normalizeSlideNumberOptions({ margin: {} }).margin).toBeUndefined();
  });

  it('detaches and deeply freezes every returned value', () => {
    const margin = { top: 1, left: 2 };
    const color = { kind: 'scheme' as const, value: 'accent1' };
    const style = { color, fontFamily: 'Aptos' };
    const input: {
      margin: { top: number; left: number };
      style: { color: typeof color; fontFamily: string };
    } = { margin, style };
    const value = normalizeSlideNumberOptions(input);
    margin.top = 99;
    color.value = 'accent2';
    style.fontFamily = 'Changed';
    expect(value.margin).toEqual({ top: 1, left: 2 });
    expect(value.style.fontFamily).toBe('Aptos');
    expect(value.style.color).toEqual({ kind: 'scheme', value: 'accent1' });
    expect(Object.isFrozen(value)).toBe(true);
    expect(Object.isFrozen(value.margin)).toBe(true);
    expect(Object.isFrozen(value.style)).toBe(true);
    expect(Object.isFrozen(value.style.color)).toBe(true);
  });

  it('accepts null-prototype data objects without invoking accessors', () => {
    const input = Object.create(null) as Record<string, unknown>;
    input.x = 1;
    input.style = Object.assign(Object.create(null), { bold: true });
    expect(normalizeSlideNumberOptions(input)).toMatchObject({
      x: 1,
      style: { bold: true },
    });

    let getterCalls = 0;
    const accessor = Object.create(null) as Record<string, unknown>;
    Object.defineProperty(accessor, 'x', {
      enumerable: true,
      get() {
        getterCalls += 1;
        return 10;
      },
    });
    expect(() => normalizeSlideNumberOptions(accessor)).toThrow(/data property/);
    expect(getterCalls).toBe(0);

    const nestedAccessor = { style: Object.create(null) as Record<string, unknown> };
    Object.defineProperty(nestedAccessor.style, 'fontFamily', {
      enumerable: true,
      get() {
        getterCalls += 1;
        return 'Never';
      },
    });
    expect(() => normalizeSlideNumberOptions(nestedAccessor)).toThrow(/data property/);
    expect(getterCalls).toBe(0);
  });

  it('rejects unsafe descriptors, unknown fields, invalid values, and exotic nested input', () => {
    class Options {
      x = 1;
    }
    const sparse = [1, 2, 3, 4] as number[];
    delete sparse[2];
    const inherited = Object.create({ x: 1 }) as Record<string, unknown>;
    inherited.y = 2;
    const invalid: unknown[] = [
      undefined,
      null,
      [],
      'slide number',
      new Options(),
      inherited,
      { unknown: true },
      { [Symbol('unknown')]: true },
      { x: Number.NaN },
      { x: 27_273_042_316_901 },
      { y: -27_273_042_316_901 },
      { width: 0 },
      { height: -1 },
      { width: 27_273_042_316_901 },
      { align: 'distributed' },
      { rtl: 1 },
      { valign: 'center' },
      { margin: sparse },
      { margin: [1, 2, 3] },
      { margin: Object.assign(Object.create({}), { top: 1 }) },
      { margin: { top: Number.POSITIVE_INFINITY } },
      { margin: { top: 169_094 } },
      { style: [] },
      { style: { unknown: true } },
      { style: { fontFamily: '' } },
      { style: { fontFamily: 'bad\u0000font' } },
      { style: { fontSize: 0.99 } },
      { style: { fontSize: 4000.01 } },
      { style: { lang: '   ' } },
      { style: { bold: 1 } },
      { style: { italic: 'true' } },
      { style: { transparency: -0.001 } },
      { style: { transparency: 100.001 } },
      { style: { color: { kind: 'srgb', value: '#FF3399' } } },
      { style: { color: { kind: 'srgb', value: 'FFF' } } },
      { style: { color: { kind: 'scheme', value: 'accent7' } } },
      { style: { color: Object.assign(Object.create({}), { kind: 'srgb', value: 'FF3399' }) } },
    ];
    for (const value of invalid) {
      expect(() => normalizeSlideNumberOptions(value), JSON.stringify(value)).toThrow();
    }
  });
});

describe('slide number reader', () => {
  it('reads a native direct field through alternate namespace prefixes without mutation', () => {
    const source = slideXml({
      prefixP: 'q',
      prefixA: 'd',
      bodyProperties: ' anchor="ctr" lIns="12700" tIns="25400" rIns="38100" bIns="50800"',
      paragraphProperties: ' algn="just" rtl="1"',
      fieldProperties: ' lang="zh-CN" b="1" i="1" sz="1813"',
      fieldStyle: '<d:solidFill><d:srgbClr val="ff3399"><d:alpha val="75000"/></d:srgbClr></d:solidFill><d:latin typeface="Aptos"/><d:ea typeface="Aptos"/><d:cs typeface="Aptos"/>',
      cachedText: '7',
    });
    const pkg = packageWith(source);
    const before = pkg.requirePart('/ppt/slides/slide1.xml').bytes.slice();
    expect(readSlideNumber(pkg, '/ppt/slides/slide1.xml', 'slide')).toEqual({
      x: 0,
      y: 0,
      width: 800_000,
      height: 300_000,
      align: 'justify',
      rtl: true,
      valign: 'middle',
      margin: { top: 2, right: 3, bottom: 4, left: 1 },
      style: {
        fontFamily: 'Aptos',
        fontSize: 18.13,
        lang: 'zh-CN',
        bold: true,
        italic: true,
        color: { kind: 'srgb', value: 'FF3399' },
        transparency: 25,
      },
    });
    expect(pkg.requirePart('/ppt/slides/slide1.xml').bytes).toEqual(before);
  });

  it('merges PptxGenJS list defaults with field overrides and ignores cached text equality', () => {
    const first = packageWith(slideXml({
      listStyle: '<a:lvl1pPr><a:defRPr sz="2400" lang="en-US" b="0"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:latin typeface="Aptos Display"/></a:defRPr></a:lvl1pPr>',
      fieldProperties: ' lang="fr-FR" b="1"',
      cachedText: '1001',
    }));
    const second = packageWith(slideXml({
      listStyle: '<a:lvl1pPr><a:defRPr sz="2400" lang="en-US" b="0"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:latin typeface="Aptos Display"/></a:defRPr></a:lvl1pPr>',
      fieldProperties: ' lang="fr-FR" b="1"',
      cachedText: '42',
    }));
    const left = readSlideNumber(first, '/ppt/slides/slide1.xml', 'slide');
    const right = readSlideNumber(second, '/ppt/slides/slide1.xml', 'slide');
    expect(left).toEqual({
      x: 0,
      y: 0,
      width: 800_000,
      height: 300_000,
      align: 'left',
      rtl: false,
      style: {
        fontFamily: 'Aptos Display',
        fontSize: 24,
        lang: 'fr-FR',
        bold: true,
        italic: false,
        color: { kind: 'scheme', value: 'accent2' },
      },
    });
    expect(right).toEqual(left);
    expect(right).not.toBe(left);
    expect(Object.isFrozen(right)).toBe(true);
    expect(Object.isFrozen(right?.style)).toBe(true);
    expect(Object.isFrozen(right?.style.color)).toBe(true);
  });

  it('keeps explicit zero transparency and treats a direct color as a complete fill override', () => {
    const zero = packageWith(slideXml({
      fieldStyle: '<a:solidFill><a:schemeClr val="accent1"><a:alpha val="100000"/></a:schemeClr></a:solidFill>',
    }));
    expect(readSlideNumber(zero, '/ppt/slides/slide1.xml', 'slide')?.style).toMatchObject({
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 0,
    });

    const override = packageWith(slideXml({
      listStyle: '<a:lvl1pPr><a:defRPr><a:solidFill><a:schemeClr val="accent2"><a:alpha val="50000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr>',
      fieldStyle: '<a:solidFill><a:srgbClr val="FF3399"/></a:solidFill>',
    }));
    expect(readSlideNumber(override, '/ppt/slides/slide1.xml', 'slide')?.style).toMatchObject({
      color: { kind: 'srgb', value: 'FF3399' },
    });
    expect(readSlideNumber(override, '/ppt/slides/slide1.xml', 'slide')?.style.transparency)
      .toBeUndefined();
  });

  it('reads direct layout and enabled master owners but rejects a disabled master', () => {
    const layout = packageWith(slideXml({ root: 'sldLayout' }), '/ppt/slideLayouts/slideLayout1.xml');
    expect(readSlideNumber(layout, '/ppt/slideLayouts/slideLayout1.xml', 'layout')).toBeDefined();

    const enabled = packageWith(slideXml({ root: 'sldMaster', tail: '<p:hf sldNum="1"/>' }), '/ppt/slideMasters/slideMaster1.xml');
    expect(readSlideNumber(enabled, '/ppt/slideMasters/slideMaster1.xml', 'master')).toBeDefined();

    for (const token of ['0', 'false', 'off']) {
      const disabled = packageWith(
        slideXml({ root: 'sldMaster', tail: `<p:hf sldNum="${token}"/>` }),
        '/ppt/slideMasters/slideMaster1.xml',
      );
      expect(readSlideNumber(disabled, '/ppt/slideMasters/slideMaster1.xml', 'master')).toBeUndefined();
    }
  });

  it('does not guess wrong-namespace, descendant, duplicate, or malformed field state', () => {
    const valid = slideXml();
    const cases = [
      valid.replace(`xmlns:p="${P}"`, 'xmlns:p="urn:wrong"'),
      valid.replace('<p:sp><p:nvSpPr>', '<p:grpSp><p:sp><p:nvSpPr>').replace('</p:sp></p:spTree>', '</p:sp></p:grpSp></p:spTree>'),
      valid.replace('</p:spTree>', `${shapeXml()}</p:spTree>`),
      valid.replace('<p:cNvPr id="2"', '<p:cNvPr id="1"'),
      valid.replace('<p:cNvPr id="1"', '<p:cNvPr id="02"'),
      valid.replace('<p:spPr>', '<p:spPr/><p:spPr>'),
      valid.replace('<a:xfrm>', '<a:xfrm/><a:xfrm>'),
      valid.replace('<a:off x="0" y="0"/>', '<a:offMissing x="0" y="0"/>'),
      valid.replace('<a:off x="0" y="0"/>', '<a:off x="0" y="0"/><a:off x="1" y="1"/>'),
      valid.replace('cx="800000"', 'cx="0"'),
      valid.replace('<p:txBody>', '<p:txBody/><p:txBody>'),
      valid.replace('<a:p>', '<a:p/><a:p>'),
      valid.replace('<a:fld id="{F7021451-1387-4CA6-816F-3879F97B5CBC}" type="slidenum">', '<a:r><a:t>x</a:t></a:r><a:fld id="{F7021451-1387-4CA6-816F-3879F97B5CBC}" type="slidenum">'),
      valid.replace('type="slidenum"', 'type="datetime"'),
      valid.replace('<a:t>1</a:t>', '<a:t>1</a:t><a:t>2</a:t>'),
      valid.replace('<a:t>1</a:t>', '<a:t><a:ext/></a:t>'),
      valid.replace('<a:fld ', '<a:fld type="slidenum"><a:t>2</a:t></a:fld><a:fld '),
      valid.replace('<a:rPr', '<a:rPr/><a:rPr'),
      valid.replace('<a:lstStyle/>', '<a:lstStyle/><a:lstStyle/>'),
      valid.replace('<a:lstStyle/>', '<a:lstStyle><a:lvl1pPr/><a:lvl1pPr/></a:lstStyle>'),
      valid.replace('<a:pPr', '<a:pPr algn="dist"'),
      valid.replace('<a:bodyPr', '<a:bodyPr anchor="invalid"'),
      valid.replace('<a:rPr', '<a:rPr sz="99"'),
      valid.replace('<a:rPr', '<a:rPr><a:solidFill><a:srgbClr val="XYZ"/></a:solidFill></a:rPr><a:rPr'),
      valid.replace('<a:rPr', '<a:rPr><a:solidFill><a:srgbClr val="FF3399"><a:alpha val="50000"><a:ext/></a:alpha></a:srgbClr></a:solidFill></a:rPr><a:rPr'),
      valid.replace('<a:endParaRPr', '<a:br/><a:endParaRPr'),
    ];
    for (const [index, source] of cases.entries()) {
      const pkg = packageWith(source);
      const before = pkg.requirePart('/ppt/slides/slide1.xml').bytes.slice();
      expect(readSlideNumber(pkg, '/ppt/slides/slide1.xml', 'slide'), `case ${index}`).toBeUndefined();
      expect(pkg.requirePart('/ppt/slides/slide1.xml').bytes).toEqual(before);
    }
  });

  it('does not mistake ordinary fields or qualified lookalikes for direct slide numbers', () => {
    const ordinary = slideXml().replace('type="sldNum"', 'type="body"');
    expect(readSlideNumber(packageWith(ordinary), '/ppt/slides/slide1.xml', 'slide')).toBeUndefined();

    const qualified = slideXml().replace(
      '<p:ph type="sldNum"',
      '<p:ph xmlns:x="urn:foreign" x:type="sldNum" type="body"',
    );
    expect(readSlideNumber(packageWith(qualified), '/ppt/slides/slide1.xml', 'slide')).toBeUndefined();
  });
});

describe('slide number writer', () => {
  it('creates canonical XML with minimum-free ids before extLst', () => {
    const uri = '/ppt/slides/slide1.xml';
    const pkg = packageWith(blankOwnerXml({ occupiedPreferredIndex: true }));
    replaceSlideNumber(pkg, uri, 'slide', {
      x: 0,
      y: 10,
      width: 900_000,
      height: 400_000,
      align: 'justify',
      rtl: true,
      valign: 'bottom',
      margin: [1, 2, 3, 4],
      style: {
        fontFamily: 'Aptos',
        fontSize: 20,
        lang: 'zh-CN',
        bold: true,
        italic: true,
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 12.5,
      },
    }, '7');

    expect(readSlideNumber(pkg, uri, 'slide')).toEqual({
      x: 0,
      y: 10,
      width: 900_000,
      height: 400_000,
      align: 'justify',
      rtl: true,
      valign: 'bottom',
      margin: { top: 1, right: 2, bottom: 3, left: 4 },
      style: {
        fontFamily: 'Aptos',
        fontSize: 20,
        lang: 'zh-CN',
        bold: true,
        italic: true,
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 12.5,
      },
    });
    const xml = partText(pkg, uri);
    expect(xml).toContain('<p:cNvPr id="2" name="Slide Number 2"/>');
    expect(xml).toContain('<p:ph type="sldNum" sz="quarter" idx="0"/>');
    expect(xml).toContain('<a:fld id="{F7021451-1387-4CA6-816F-3879F97B5CBC}" type="slidenum">');
    expect(xml).toContain('<a:t>7</a:t>');
    expect(xml).toContain('<a:alpha val="87500"/>');
    expect(xml).toContain('<a:latin typeface="Aptos"/><a:ea typeface="Aptos"/><a:cs typeface="Aptos"/>');
    expect(xml.indexOf('name="Slide Number 2"')).toBeLessThan(xml.indexOf('<p:extLst>'));
    expect(xml).not.toContain('id="25" name="Slide Number');
  });

  it('patches supported spans and PptxGenJS defaults while preserving opaque state', () => {
    const uri = '/ppt/slides/slide1.xml';
    const source = slideXml({
      bodyProperties: ' anchor="ctr" data-body="keep"',
      paragraphProperties: ' algn="ctr" data-paragraph="keep"',
      listStyle: '<a:lvl1pPr data-level="keep"><a:defRPr sz="2400" lang="en-US" b="1" data-default="keep"><a:solidFill><a:schemeClr val="accent2"/></a:solidFill><a:latin typeface="Old"/><a:extLst><a:ext uri="default-opaque"/></a:extLst></a:defRPr></a:lvl1pPr>',
      fieldProperties: ' lang="en-US" b="0" data-field="keep"',
      fieldStyle: '<a:extLst><a:ext uri="field-opaque"/></a:extLst>',
      cachedText: '3',
    })
      .replace('<a:off x="0" y="0"/>', '<a:off x="0" y="0" data-off="keep"/>')
      .replace('<a:ext cx="800000" cy="300000"/>', '<a:ext cx="800000" cy="300000" data-ext="keep"/>')
      .replace('<p:cNvPr id="2" name="Slide Number 2"/>', '<p:cNvPr id="2" name="Imported Number" data-id="keep"/>')
      .replace('<p:ph type="sldNum" sz="quarter" idx="4294967295"/>', '<p:ph type="sldNum" sz="quarter" idx="42" data-ph="keep"/>')
      .replace('</p:sp>', '<p:extLst><p:ext uri="shape-opaque"/></p:extLst></p:sp>');
    const pkg = packageWith(source);
    replaceSlideNumber(pkg, uri, 'slide', {
      x: 100,
      y: 200,
      width: 300,
      height: 400,
      align: 'right',
      rtl: true,
      margin: { left: 5 },
      style: {
        fontFamily: 'New',
        fontSize: 12,
        lang: 'fr-FR',
        italic: true,
        color: { kind: 'srgb', value: '00FF00' },
      },
    }, '9');

    const output = partText(pkg, uri);
    for (const fragment of [
      'name="Imported Number" data-id="keep"',
      'idx="42" data-ph="keep"',
      'data-off="keep"',
      'data-ext="keep"',
      'data-body="keep"',
      'data-paragraph="keep"',
      'data-default="keep"',
      'data-field="keep"',
      'uri="default-opaque"',
      'uri="field-opaque"',
      'uri="shape-opaque"',
    ]) expect(output).toContain(fragment);
    expect(output).not.toContain('typeface="Old"');
    expect(readSlideNumber(pkg, uri, 'slide')).toMatchObject({
      x: 100,
      y: 200,
      width: 300,
      height: 400,
      align: 'right',
      rtl: true,
      margin: { left: 5 },
      style: {
        fontFamily: 'New',
        fontSize: 12,
        lang: 'fr-FR',
        bold: false,
        italic: true,
        color: { kind: 'srgb', value: '00FF00' },
      },
    });
    expect(output).toContain('<a:t>9</a:t>');
    expect(output).toContain(`id="{F7021451-1387-4CA6-816F-3879F97B5CBC}" type="slidenum"`);
  });

  it('keeps same values and absent clears exact, updates only cache, and clears a unique field', () => {
    const uri = '/ppt/slides/slide1.xml';
    const pkg = packageWith(slideXml());
    const value = readSlideNumber(pkg, uri, 'slide');
    if (!value) throw new Error('Expected slide number');
    const before = packageSnapshot(pkg);
    replaceSlideNumber(pkg, uri, 'slide', value, '1');
    expect(packageSnapshot(pkg)).toEqual(before);

    expect(replaceSlideNumberCachedText(pkg, uri, '2')).toBe(true);
    const afterCache = partText(pkg, uri);
    expect(afterCache).toBe(partText(packageWith(slideXml({ cachedText: '2' })), uri));
    expect(replaceSlideNumberCachedText(pkg, uri, '2')).toBe(false);

    replaceSlideNumber(pkg, uri, 'slide', undefined, '2');
    expect(readSlideNumber(pkg, uri, 'slide')).toBeUndefined();
    expect(partText(pkg, uri)).not.toContain('type="sldNum"');
    const cleared = packageSnapshot(pkg);
    replaceSlideNumber(pkg, uri, 'slide', undefined, '2');
    expect(packageSnapshot(pkg)).toEqual(cleared);
  });

  it('canonicalizes one opaque placeholder but rejects ambiguous identity before mutation', () => {
    const uri = '/ppt/slides/slide1.xml';
    const opaque = packageWith(slideXml().replace('type="slidenum"', 'type="datetime"'));
    expect(readSlideNumber(opaque, uri, 'slide')).toBeUndefined();
    replaceSlideNumber(opaque, uri, 'slide', { align: 'center' }, '4');
    expect(readSlideNumber(opaque, uri, 'slide')).toMatchObject({ align: 'center' });
    expect(partText(opaque, uri)).toContain('id="2" name="Slide Number 2"');
    expect(partText(opaque, uri)).toContain('idx="4294967295"');

    const ambiguous = packageWith(
      slideXml().replace('type="slidenum"', 'type="datetime"').replace(
        '<p:cNvPr id="1"',
        '<p:cNvPr id="2"',
      ),
    );
    const before = packageSnapshot(ambiguous);
    expect(() => replaceSlideNumber(ambiguous, uri, 'slide', {}, '1')).toThrow(/identity|ambiguous/i);
    expect(packageSnapshot(ambiguous)).toEqual(before);
  });

  it('rejects multiple direct placeholders for assignment and clear without mutation', () => {
    const uri = '/ppt/slides/slide1.xml';
    const source = slideXml().replace(
      '</p:spTree>',
      `<p:sp>${shapeXml().replace('id="2"', 'id="3"').replace('idx="4294967295"', 'idx="8"')}</p:sp></p:spTree>`,
    );
    const pkg = packageWith(source);
    const before = packageSnapshot(pkg);
    expect(() => replaceSlideNumber(pkg, uri, 'slide', {}, '1')).toThrow(/ambiguous/i);
    expect(packageSnapshot(pkg)).toEqual(before);
    expect(() => replaceSlideNumber(pkg, uri, 'slide', undefined, '1')).toThrow(/ambiguous/i);
    expect(packageSnapshot(pkg)).toEqual(before);
  });

  it('enables, inserts, preserves, and disables the master header-footer flag in schema order', () => {
    const uri = '/ppt/slideMasters/slideMaster1.xml';
    const pkg = packageWith(blankOwnerXml({
      root: 'sldMaster',
      rootTail: '<p:sldLayoutIdLst/><p:txStyles data-styles="keep"/>',
    }), uri);
    replaceSlideNumber(pkg, uri, 'master', {}, '‹#›');
    let output = partText(pkg, uri);
    expect(output).toContain('<p:hf sldNum="1"/>');
    expect(output.indexOf('<p:sldLayoutIdLst/>')).toBeLessThan(output.indexOf('<p:hf'));
    expect(output.indexOf('<p:hf')).toBeLessThan(output.indexOf('<p:txStyles'));

    output = output.replace('<p:hf sldNum="1"/>', '<p:hf hdr="1" ftr="0" data-hf="keep" sldNum="1"/>');
    pkg.setPart(uri, output, SLIDE_CONTENT_TYPE);
    replaceSlideNumber(pkg, uri, 'master', undefined, '‹#›');
    output = partText(pkg, uri);
    expect(output).toContain('<p:hf hdr="1" ftr="0" data-hf="keep" sldNum="0"/>');
    expect(output).not.toContain('type="sldNum"');
  });

  it('validates before package access and rolls every edit back with an outer transaction', () => {
    const empty = OpcPackage.create();
    expect(() => replaceSlideNumber(
      empty,
      '/missing.xml',
      'slide',
      { width: 0 },
      '1',
    )).toThrow(/width/i);

    const uri = '/ppt/slides/slide1.xml';
    const pkg = packageWith(blankOwnerXml());
    const before = packageSnapshot(pkg);
    expect(() => pkg.transaction(() => {
      replaceSlideNumber(pkg, uri, 'slide', { style: { bold: true } }, '1');
      throw new Error('injected outer failure');
    })).toThrow('injected outer failure');
    expect(packageSnapshot(pkg)).toEqual(before);
  });
});

interface SlideXmlOptions {
  readonly root?: 'sld' | 'sldLayout' | 'sldMaster';
  readonly prefixP?: string;
  readonly prefixA?: string;
  readonly bodyProperties?: string;
  readonly paragraphProperties?: string;
  readonly fieldProperties?: string;
  readonly fieldStyle?: string;
  readonly listStyle?: string;
  readonly cachedText?: string;
  readonly tail?: string;
}

function slideXml(options: SlideXmlOptions = {}): string {
  const p = options.prefixP ?? 'p';
  const a = options.prefixA ?? 'a';
  const root = options.root ?? 'sld';
  const listStyle = options.listStyle === undefined
    ? `<${a}:lstStyle/>`
    : `<${a}:lstStyle>${rewritePrefixes(options.listStyle, a)}</${a}:lstStyle>`;
  const fieldStyle = rewritePrefixes(options.fieldStyle ?? '', a);
  return `<?xml version="1.0"?><${p}:${root} xmlns:${p}="${P}" xmlns:${a}="${A}"><${p}:cSld><${p}:spTree><${p}:nvGrpSpPr><${p}:cNvPr id="1" name=""/><${p}:cNvGrpSpPr/><${p}:nvPr/></${p}:nvGrpSpPr><${p}:grpSpPr/><${p}:sp>${shapeXml({
    p,
    a,
    ...(options.bodyProperties === undefined ? {} : { bodyProperties: options.bodyProperties }),
    ...(options.paragraphProperties === undefined ? {} : { paragraphProperties: options.paragraphProperties }),
    ...(options.fieldProperties === undefined ? {} : { fieldProperties: options.fieldProperties }),
    fieldStyle,
    listStyle,
    ...(options.cachedText === undefined ? {} : { cachedText: options.cachedText }),
  })}</${p}:sp></${p}:spTree></${p}:cSld>${options.tail ?? ''}</${p}:${root}>`;
}

function shapeXml(options: {
  readonly p?: string;
  readonly a?: string;
  readonly bodyProperties?: string;
  readonly paragraphProperties?: string;
  readonly fieldProperties?: string;
  readonly fieldStyle?: string;
  readonly listStyle?: string;
  readonly cachedText?: string;
} = {}): string {
  const p = options.p ?? 'p';
  const a = options.a ?? 'a';
  return `<${p}:nvSpPr><${p}:cNvPr id="2" name="Slide Number 2"/><${p}:cNvSpPr/><${p}:nvPr><${p}:ph type="sldNum" sz="quarter" idx="4294967295"/></${p}:nvPr></${p}:nvSpPr><${p}:spPr><${a}:xfrm><${a}:off x="0" y="0"/><${a}:ext cx="800000" cy="300000"/></${a}:xfrm></${p}:spPr><${p}:txBody><${a}:bodyPr${options.bodyProperties ?? ''}/>${options.listStyle ?? `<${a}:lstStyle/>`}<${a}:p><${a}:pPr${options.paragraphProperties ?? ''}/><${a}:fld id="{F7021451-1387-4CA6-816F-3879F97B5CBC}" type="slidenum"><${a}:rPr${options.fieldProperties ?? ''}>${options.fieldStyle ?? ''}</${a}:rPr><${a}:t>${options.cachedText ?? '1'}</${a}:t></${a}:fld><${a}:endParaRPr lang="en-US"/></${a}:p></${p}:txBody>`;
}

function rewritePrefixes(value: string, drawingPrefix: string): string {
  return value.replaceAll('a:', `${drawingPrefix}:`);
}

function packageWith(
  xml: string,
  uri = '/ppt/slides/slide1.xml',
): OpcPackage {
  const pkg = OpcPackage.create();
  pkg.setPart(uri, xml, SLIDE_CONTENT_TYPE);
  return pkg;
}

function blankOwnerXml(options: {
  readonly root?: 'sld' | 'sldLayout' | 'sldMaster';
  readonly occupiedPreferredIndex?: boolean;
  readonly rootTail?: string;
} = {}): string {
  const occupied = options.occupiedPreferredIndex
    ? '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="4294967295"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>'
    : '';
  return `<?xml version="1.0"?><p:${options.root ?? 'sld'} xmlns:p="${P}" xmlns:a="${A}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${occupied}<p:extLst><p:ext uri="tree-opaque"/></p:extLst></p:spTree></p:cSld>${options.rootTail ?? ''}</p:${options.root ?? 'sld'}>`;
}

function partText(pkg: OpcPackage, uri: string): string {
  return new TextDecoder().decode(pkg.requirePart(uri).bytes);
}

function packageSnapshot(pkg: OpcPackage): unknown {
  return {
    parts: pkg.parts.map(({ uri, contentType, bytes }) => ({
      uri,
      contentType,
      bytes: [...bytes],
    })),
    relationships: pkg.parts.map(({ uri }) => ({
      uri,
      relationships: pkg.relationships(uri).map((relationship) => ({ ...relationship })),
    })),
    mutations: pkg.mutations.map((mutation) => ({ ...mutation })),
  };
}

const _compileOnly: SlideNumberOptions = {
  margin: { top: 1 },
  style: { color: { kind: 'scheme', value: 'accent1' } },
};
void _compileOnly;
