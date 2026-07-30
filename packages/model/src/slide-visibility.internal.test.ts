import { describe, expect, it } from 'vitest';
import {
  LosslessXmlDocument,
  LosslessXmlError,
} from '@pptx/lossless-xml';
import { ModelParseError } from './errors.js';
import {
  readSlideHidden,
  replaceSlideHidden,
} from './slide-visibility.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';

function slideXml(
  attributes = '',
  body = '<p:cSld/>',
  prefix = 'p',
): string {
  const qualifier = prefix === '' ? '' : `${prefix}:`;
  const declaration = prefix === ''
    ? `xmlns="${PRESENTATION_NAMESPACE}"`
    : `xmlns:${prefix}="${PRESENTATION_NAMESPACE}"`;
  return `<${qualifier}sld ${declaration}${attributes}>${body}</${qualifier}sld>`;
}

function expectRejectedWithoutPatch(source: string, value: boolean): void {
  const xml = LosslessXmlDocument.parse(source);
  expect(() => replaceSlideHidden(xml, value)).toThrow(ModelParseError);
  expect(xml.changed).toBe(false);
  expect(xml.serialize()).toBe(source);
}

describe('slide visibility codec', () => {
  it('reads absent and all six legal tokens without changing source state', () => {
    const cases: readonly (readonly [string, boolean])[] = [
      ['', false],
      [' show="0"', true],
      [' show="false"', true],
      [' show="off"', true],
      [' show="1"', false],
      [' show="true"', false],
      [' show="on"', false],
    ];
    for (const [attributes, expected] of cases) {
      const source = slideXml(attributes);
      const xml = LosslessXmlDocument.parse(source);
      expect(readSlideHidden(xml), source).toBe(expected);
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });

  it('accepts alternate prefixes and ignores qualified and descendant lookalikes', () => {
    const source =
      `<?xml version="1.0"?>\n<!--before-->\n` +
      slideXml(
        ' data="KEEP" xmlns:x="urn:foreign" x:show="0"',
        '\n  <!--inside-->\n  <q:cSld data="KEEP"><q:sp show="0"/></q:cSld>\n',
        'q',
      ) +
      '\n<!--after-->';
    const xml = LosslessXmlDocument.parse(source);
    expect(readSlideHidden(xml)).toBe(false);
    expect(xml.changed).toBe(false);
    expect(xml.serialize()).toBe(source);
  });

  it('returns undefined for unsafe ownership and unknown lexical state', () => {
    const cases = [
      '<p:sld xmlns:p="urn:wrong"/>',
      slideXml(' show=""'),
      slideXml(' show="maybe"'),
      slideXml(' show="0" show="1"'),
      `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
      `${slideXml()}${slideXml()}`,
      `<p:notSlide xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
    ];
    for (const source of cases) {
      const xml = LosslessXmlDocument.parse(source);
      expect(readSlideHidden(xml), source).toBeUndefined();
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
    expect(() => LosslessXmlDocument.parse('<p:sld>')).toThrow(LosslessXmlError);
  });

  it('inserts, canonicalizes, removes, and repairs only direct unqualified show', () => {
    const visibleSource = slideXml(
      ' custom="KEEP" xmlns:x="urn:foreign" x:show="KEEP"',
      '<p:cSld><p:sp show="KEEP"/></p:cSld>',
    );
    const visible = LosslessXmlDocument.parse(visibleSource);
    expect(replaceSlideHidden(visible, true)).toBe(true);
    expect(visible.serialize()).toBe(visibleSource.replace('><p:cSld>', ' show="0"><p:cSld>'));

    for (const token of ['false', 'off', 'maybe']) {
      const source = slideXml(` custom="KEEP" show='${token}' tail="KEEP"`);
      const xml = LosslessXmlDocument.parse(source);
      expect(replaceSlideHidden(xml, true), token).toBe(true);
      expect(xml.serialize()).toBe(source.replace(token, '0'));
    }

    for (const token of ['0', 'false', 'off', '1', 'true', 'on', 'maybe']) {
      const source = slideXml(` custom="KEEP"  show="${token}" tail="KEEP"`);
      const xml = LosslessXmlDocument.parse(source);
      expect(replaceSlideHidden(xml, false), token).toBe(true);
      expect(xml.serialize()).toBe(slideXml(' custom="KEEP" tail="KEEP"'));
    }

    const multiline = slideXml('\n\tshow="0"\n\tcustom="KEEP"');
    const multilineXml = LosslessXmlDocument.parse(multiline);
    expect(replaceSlideHidden(multilineXml, false)).toBe(true);
    expect(multilineXml.serialize()).toBe(slideXml('\n\n\tcustom="KEEP"'));
  });

  it('patches self-closing default-namespace roots without expanding them', () => {
    const visibleSource = `<sld xmlns="${PRESENTATION_NAMESPACE}"/>`;
    const visible = LosslessXmlDocument.parse(visibleSource);
    expect(readSlideHidden(visible)).toBe(false);
    expect(replaceSlideHidden(visible, true)).toBe(true);
    const hiddenSource = `<sld xmlns="${PRESENTATION_NAMESPACE}" show="0"/>`;
    expect(visible.serialize()).toBe(hiddenSource);

    const hidden = LosslessXmlDocument.parse(hiddenSource);
    expect(readSlideHidden(hidden)).toBe(true);
    expect(replaceSlideHidden(hidden, false)).toBe(true);
    expect(hidden.serialize()).toBe(visibleSource);
  });

  it('preserves canonical no-ops without creating patches', () => {
    const visibleSource = slideXml(' custom="KEEP"');
    const visible = LosslessXmlDocument.parse(visibleSource);
    expect(replaceSlideHidden(visible, false)).toBe(false);
    expect(visible.changed).toBe(false);
    expect(visible.serialize()).toBe(visibleSource);

    const hiddenSource = slideXml(' custom="KEEP" show="0"');
    const hidden = LosslessXmlDocument.parse(hiddenSource);
    expect(replaceSlideHidden(hidden, true)).toBe(false);
    expect(hidden.changed).toBe(false);
    expect(hidden.serialize()).toBe(hiddenSource);
  });

  it('rejects ambiguous roots before making any patch', () => {
    const unsafe = [
      '<p:sld xmlns:p="urn:wrong"/>',
      slideXml(' show="0" show="1"'),
      `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
      `${slideXml()}${slideXml()}`,
      `<p:notSlide xmlns:p="${PRESENTATION_NAMESPACE}"/>`,
    ];
    for (const source of unsafe) {
      expectRejectedWithoutPatch(source, true);
      expectRejectedWithoutPatch(source, false);
    }
  });
});
