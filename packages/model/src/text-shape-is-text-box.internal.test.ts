import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  readTextShapeIsTextBox,
  replaceTextShapeIsTextBox,
} from './text-shape-is-text-box.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const PART_URI = '/ppt/slides/slide1.xml';

function fixture(
  attributes = '',
  prefix = 'p',
  nonVisualBody = '',
): string {
  const qualified = prefix === '' ? '' : `${prefix}:`;
  const declaration = prefix === ''
    ? `xmlns="${PRESENTATION_NAMESPACE}"`
    : `xmlns:${prefix}="${PRESENTATION_NAMESPACE}"`;
  return `<?xml version="1.0"?>\n<!--before-->\n<${qualified}sp ${declaration} data="SHAPE">`
    + `<${qualified}nvSpPr data="NV"><${qualified}cNvPr id="2" name="Keep"/>`
    + `<${qualified}cNvSpPr data="KEEP"${attributes}>${nonVisualBody}`
    + `</${qualified}cNvSpPr><${qualified}nvPr/></${qualified}nvSpPr>`
    + `<${qualified}spPr/><${qualified}txBody/></${qualified}sp>\n<!--after-->`;
}

function selfClosingFixture(attributes = '', prefix = 'p'): string {
  const qualified = prefix === '' ? '' : `${prefix}:`;
  return fixture(attributes, prefix).replace(
    `<${qualified}cNvSpPr data="KEEP"${attributes}></${qualified}cNvSpPr>`,
    `<${qualified}cNvSpPr data="KEEP"${attributes}/>`,
  );
}

function parse(source: string): {
  readonly xml: LosslessXmlDocument;
  readonly shape: XmlElement;
} {
  const xml = LosslessXmlDocument.parse(source);
  const shape = xml.roots.find(({ localName }) => localName === 'sp');
  if (!shape) throw new Error('Fixture shape is missing');
  return { xml, shape };
}

function expectRejectedWithoutPatch(source: string, value: boolean): void {
  const { xml, shape } = parse(source);
  expect(() => replaceTextShapeIsTextBox(xml, shape, value, PART_URI))
    .toThrow(ModelParseError);
  expect(xml.changed).toBe(false);
  expect(xml.serialize()).toBe(source);
}

describe('text shape isTextBox codec', () => {
  it('reads absent and all six legal tokens without changing source bytes', () => {
    const cases: readonly (readonly [string, boolean])[] = [
      ['', false],
      [' txBox="1"', true],
      [' txBox="true"', true],
      [' txBox="on"', true],
      [' txBox="0"', false],
      [' txBox="false"', false],
      [' txBox="off"', false],
    ];
    for (const [attributes, expected] of cases) {
      const source = fixture(attributes);
      const { xml, shape } = parse(source);
      expect(readTextShapeIsTextBox(xml, shape), source).toBe(expected);
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });

  it('accepts an alternate presentation prefix and ignores unrelated descendants', () => {
    const source = fixture(
      ' xmlns:x="urn:foreign" x:keep="yes" txBox="true"',
      'q',
      '<q:extLst><q:cNvSpPr txBox="off"/></q:extLst>',
    );
    const { xml, shape } = parse(source);
    expect(readTextShapeIsTextBox(xml, shape)).toBe(true);
    expect(xml.changed).toBe(false);
    expect(xml.serialize()).toBe(source);
  });

  it('returns undefined for unsafe ownership, attributes, and lexical values', () => {
    const valid = fixture();
    const cases = [
      fixture(' txBox=""'),
      fixture(' txBox="maybe"'),
      fixture(' txBox="1" txBox="0"'),
      fixture(' xmlns:x="urn:foreign" x:txBox="1"'),
      valid.replace(PRESENTATION_NAMESPACE, 'urn:wrong'),
      valid.replace(
        '<p:nvSpPr data="NV">',
        '<p:nvSpPr><p:cNvSpPr/></p:nvSpPr><p:nvSpPr data="NV">',
      ),
      valid.replace('<p:cNvSpPr data="KEEP"></p:cNvSpPr>', ''),
      valid.replace(
        '<p:cNvSpPr data="KEEP"></p:cNvSpPr>',
        '<p:cNvSpPr/><p:cNvSpPr/>',
      ),
      valid.replace(
        '<p:cNvSpPr data="KEEP"></p:cNvSpPr>',
        '<x:cNvSpPr xmlns:x="urn:foreign" txBox="1"/>',
      ),
      valid.replace(
        '<p:cNvSpPr data="KEEP"></p:cNvSpPr>',
        '<p:ext><p:cNvSpPr txBox="1"/></p:ext>',
      ),
    ];
    for (const source of cases) {
      const { xml, shape } = parse(source);
      expect(readTextShapeIsTextBox(xml, shape), source).toBeUndefined();
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });

  it('inserts canonical true into self-closing and expanded owners', () => {
    for (const source of [selfClosingFixture(), fixture()]) {
      const { xml, shape } = parse(source);
      expect(replaceTextShapeIsTextBox(xml, shape, true, PART_URI)).toBe(true);
      expect(xml.serialize()).toBe(source.replace(' data="KEEP"', ' data="KEEP" txBox="1"'));
    }
  });

  it('canonicalizes legal aliases and a single malformed token', () => {
    for (const token of ['true', 'on', '0', 'false', 'off', 'maybe']) {
      const source = fixture(` custom="A" txBox='${token}' tail="B"`);
      const { xml, shape } = parse(source);
      expect(replaceTextShapeIsTextBox(xml, shape, true, PART_URI), token).toBe(true);
      expect(xml.serialize()).toBe(source.replace(`txBox='${token}'`, "txBox='1'"));
    }
  });

  it('removes every single explicit token while preserving surrounding bytes', () => {
    for (const token of ['1', 'true', 'on', '0', 'false', 'off', 'maybe']) {
      const source = fixture(` custom="A"  txBox="${token}" tail="B"`);
      const { xml, shape } = parse(source);
      expect(replaceTextShapeIsTextBox(xml, shape, false, PART_URI), token).toBe(true);
      expect(xml.serialize()).toBe(fixture(' custom="A" tail="B"'));
    }

    const multiline = fixture('\n\ttxBox="1"\n\tcustom="KEEP"');
    const { xml, shape } = parse(multiline);
    expect(replaceTextShapeIsTextBox(xml, shape, false, PART_URI)).toBe(true);
    expect(xml.serialize()).toBe(fixture('\n\n\tcustom="KEEP"'));
  });

  it('preserves exact canonical no-ops without creating patches', () => {
    for (const [source, value] of [
      [fixture(), false],
      [fixture(' txBox="1"'), true],
    ] as const) {
      const { xml, shape } = parse(source);
      expect(replaceTextShapeIsTextBox(xml, shape, value, PART_URI)).toBe(false);
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });

  it('rejects ambiguous owners and attributes before making any patch', () => {
    const valid = fixture();
    const unsafe = [
      fixture(' txBox="1" txBox="0"'),
      fixture(' xmlns:x="urn:foreign" x:txBox="1"'),
      valid.replace(PRESENTATION_NAMESPACE, 'urn:wrong'),
      valid.replace('<p:cNvSpPr data="KEEP"></p:cNvSpPr>', ''),
      valid.replace(
        '<p:cNvSpPr data="KEEP"></p:cNvSpPr>',
        '<p:cNvSpPr/><p:cNvSpPr/>',
      ),
    ];
    for (const source of unsafe) {
      expectRejectedWithoutPatch(source, true);
      expectRejectedWithoutPatch(source, false);
    }
  });
});
