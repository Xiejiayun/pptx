import { LosslessXmlDocument } from '@pptx/lossless-xml';
import { describe, expect, it } from 'vitest';
import { ModelParseError } from './errors.js';
import {
  readShapeFill,
  replaceShapeFill,
} from './shape-fill.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const PART_URI = '/ppt/slides/slide1.xml';

function fixture(
  properties: string,
  options: {
    readonly rootName?: string;
    readonly presentationNamespace?: string;
    readonly drawingNamespace?: string;
  } = {},
): string {
  const rootName = options.rootName ?? 'p:sp';
  return `<${rootName} xmlns:p="${options.presentationNamespace ?? PRESENTATION_NAMESPACE}" ` +
    `xmlns:a="${options.drawingNamespace ?? DRAWING_NAMESPACE}">` +
    '<p:nvSpPr><p:cNvPr id="7" name="Keep"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    `${properties}<p:txBody><a:bodyPr/><a:p><a:r><a:t>KEEP</a:t></a:r></a:p></p:txBody>` +
    `</${rootName}>`;
}

function properties(fill = '<a:noFill/>', geometry = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>') {
  return '<p:spPr keep="PROPERTIES"><a:xfrm><a:off x="1" y="2"/>' +
    `<a:ext cx="3" cy="4"/></a:xfrm>${geometry}${fill}` +
    '<a:ln w="9"/><a:effectLst/><a:scene3d/><a:sp3d/>' +
    '<a:extLst><a:ext uri="urn:keep"><x:opaque xmlns:x="urn:test"/></a:ext></a:extLst>' +
    '</p:spPr>';
}

function parse(source: string) {
  const xml = LosslessXmlDocument.parse(source);
  const shape = xml.roots[0];
  if (!shape) throw new Error('Fixture has no shape');
  return { xml, shape };
}

describe('shape fill reader', () => {
  it('reads detached direct none, sRGB, scheme, alpha, and alternate-prefix fills', () => {
    const cases = [
      {
        source: fixture(properties('<a:noFill/>')),
        expected: { kind: 'none' },
      },
      {
        source: fixture(properties(
          '<a:solidFill><a:srgbClr val="ff0000"/></a:solidFill>',
        )),
        expected: { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } },
      },
      {
        source: fixture(properties(
          '<a:solidFill><a:schemeClr val="accent2"><a:alpha val="75000"/>' +
          '</a:schemeClr></a:solidFill>',
        )),
        expected: {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 25,
        },
      },
      {
        source:
          `<q:sp xmlns:q="${PRESENTATION_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}">` +
          '<q:spPr><d:prstGeom prst="ellipse"><d:avLst/></d:prstGeom>' +
          '<d:solidFill><d:srgbClr val="112233"><d:alpha val="49445"/>' +
          '</d:srgbClr></d:solidFill></q:spPr></q:sp>',
        expected: {
          kind: 'solid',
          color: { kind: 'srgb', value: '112233' },
          transparency: 50.555,
        },
      },
    ];

    for (const { source, expected } of cases) {
      const { xml, shape } = parse(source);
      const first = readShapeFill(xml, shape);
      const second = readShapeFill(xml, shape);
      expect(first, source).toEqual(expected);
      expect(second, source).toEqual(expected);
      expect(first).not.toBe(second);
      if (first?.kind === 'solid' && second?.kind === 'solid') {
        expect(first.color).not.toBe(second.color);
      }
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });

  it('returns undefined for absent, unsupported, nested, ambiguous, or unsafe containers', () => {
    const unsupported = [
      '<a:gradFill><a:gsLst/></a:gradFill>',
      '<a:blipFill/>',
      '<a:pattFill/>',
      '<a:grpFill/>',
    ];
    const sources = [
      fixture(properties('')),
      ...unsupported.map((fill) => fixture(properties(fill))),
      fixture('<p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
        '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>' +
        '</p:spPr>'),
      fixture('<p:spPr/><p:spPr/>'),
      fixture('<x:spPr xmlns:x="urn:wrong"><a:noFill/></x:spPr>'),
      fixture(properties('<a:noFill/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')),
      fixture(properties('<x:noFill xmlns:x="urn:wrong"/>')),
      fixture(properties('<a:noFill/>'), { rootName: 'p:pic' }),
      fixture(properties('<a:noFill/>'), { presentationNamespace: 'urn:wrong' }),
    ];
    for (const source of sources) {
      const { xml, shape } = parse(source);
      expect(readShapeFill(xml, shape), source).toBeUndefined();
      expect(xml.changed, source).toBe(false);
      expect(xml.serialize(), source).toBe(source);
    }
  });

  it('returns undefined for malformed simple choices and namespace lookalikes', () => {
    const fills = [
      '<a:noFill custom="x"/>',
      '<a:noFill><a:ext/></a:noFill>',
      '<a:solidFill/>',
      '<a:solidFill custom="x"><a:srgbClr val="FF0000"/></a:solidFill>',
      '<a:solidFill><a:srgbClr val="FFF"/></a:solidFill>',
      '<a:solidFill><a:schemeClr val="unknown"/></a:solidFill>',
      '<a:solidFill><a:srgbClr val="FF0000"/><a:schemeClr val="accent1"/></a:solidFill>',
      '<a:solidFill><a:srgbClr val="FF0000"><a:tint val="50000"/></a:srgbClr></a:solidFill>',
      '<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="100001"/></a:srgbClr></a:solidFill>',
      '<a:solidFill><x:srgbClr xmlns:x="urn:wrong" val="FF0000"/></a:solidFill>',
      `<a:solidFill><x:srgbClr xmlns:x="${DRAWING_NAMESPACE}" val="FF0000"/></a:solidFill>`,
      '<a:solidFill><a:srgbClr xmlns:a="urn:wrong" val="FF0000"/></a:solidFill>',
    ];
    for (const fill of fills) {
      const source = fixture(properties(fill));
      const { xml, shape } = parse(source);
      expect(readShapeFill(xml, shape), fill).toBeUndefined();
      expect(xml.serialize()).toBe(source);
    }
  });
});

describe('shape fill replacement', () => {
  it('preserves exact bytes for same-value assignments', () => {
    const sources = [
      fixture(properties('<a:noFill/>')),
      fixture(properties('<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')),
      fixture(properties(
        '<a:solidFill><a:schemeClr val="accent2"><a:alpha val="75000"/>' +
        '</a:schemeClr></a:solidFill>',
      )),
    ];
    const values = [
      { kind: 'none' } as const,
      { kind: 'solid', color: { kind: 'srgb', value: 'FF0000' } } as const,
      {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent2' },
        transparency: 25,
      } as const,
    ];
    for (let index = 0; index < sources.length; index += 1) {
      const source = sources[index]!;
      const { xml, shape } = parse(source);
      expect(replaceShapeFill(xml, shape, values[index]!, PART_URI)).toBe(false);
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });

  it('whole-replaces and clears every unique direct fill choice in isolation', () => {
    const choices = [
      '<a:noFill/>',
      '<a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>',
      '<a:gradFill><a:gsLst/></a:gradFill>',
      '<a:blipFill><a:blip r:embed="rId7" xmlns:r="urn:r"/></a:blipFill>',
      '<a:pattFill prst="pct10"/>',
      '<a:grpFill/>',
    ];
    for (const choice of choices) {
      const source = fixture(properties(choice));
      const replaced = parse(source);
      expect(replaceShapeFill(replaced.xml, replaced.shape, {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent4' },
        transparency: 40,
      }, PART_URI), choice).toBe(true);
      const updated = replaced.xml.serialize();
      expect(updated).toContain(
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
        '<a:solidFill><a:schemeClr val="accent4"><a:alpha val="60000"/>' +
        '</a:schemeClr></a:solidFill><a:ln w="9"/>',
      );
      expect(updated).toContain('<p:cNvPr id="7" name="Keep"/>');
      expect(updated).toContain('<a:off x="1" y="2"/><a:ext cx="3" cy="4"/>');
      expect(updated).toContain('<a:effectLst/><a:scene3d/><a:sp3d/>');
      expect(updated).toContain('<x:opaque xmlns:x="urn:test"/>');
      expect(updated).toContain('<a:t>KEEP</a:t>');

      const cleared = parse(source);
      expect(replaceShapeFill(cleared.xml, cleared.shape, undefined, PART_URI), choice)
        .toBe(true);
      const clearedXml = cleared.xml.serialize();
      expect(clearedXml).toContain(
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="9"/>',
      );
      const reparsed = parse(clearedXml);
      expect(readShapeFill(reparsed.xml, reparsed.shape)).toBeUndefined();
    }
  });

  it('inserts after unique preset or custom geometry and before later properties', () => {
    for (const geometry of [
      '<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 7"/></a:avLst></a:prstGeom>',
      '<a:custGeom><a:avLst/><a:pathLst><a:path w="1" h="1"/></a:pathLst></a:custGeom>',
    ]) {
      const source = fixture(properties('', geometry));
      const { xml, shape } = parse(source);
      expect(replaceShapeFill(xml, shape, { kind: 'none' }, PART_URI)).toBe(true);
      const updated = xml.serialize();
      expect(updated.indexOf(geometry)).toBeLessThan(updated.indexOf('<a:noFill/>'));
      expect(updated.indexOf('<a:noFill/>')).toBeLessThan(updated.indexOf('<a:ln w="9"/>'));
      expect(updated).toContain(geometry);
      expect(updated).toContain('<x:opaque xmlns:x="urn:test"/>');
    }
  });

  it('retains or adds the namespace binding needed by replacement and insertion', () => {
    const alternate =
      `<q:sp xmlns:q="${PRESENTATION_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}">` +
      '<q:spPr><d:prstGeom prst="rect"><d:avLst/></d:prstGeom><d:grpFill/>' +
      '<d:ln/></q:spPr></q:sp>';
    const alternateParsed = parse(alternate);
    expect(replaceShapeFill(alternateParsed.xml, alternateParsed.shape, {
      kind: 'solid',
      color: { kind: 'srgb', value: '123456' },
    }, PART_URI)).toBe(true);
    expect(alternateParsed.xml.serialize()).toContain(
      '<d:solidFill><d:srgbClr val="123456"/></d:solidFill><d:ln/>',
    );

    const localGeometry =
      `<p:sp xmlns:p="${PRESENTATION_NAMESPACE}"><p:spPr>` +
      `<d:prstGeom xmlns:d="${DRAWING_NAMESPACE}" prst="rect"><d:avLst/></d:prstGeom>` +
      '<p:extLst/></p:spPr></p:sp>';
    const localParsed = parse(localGeometry);
    expect(replaceShapeFill(localParsed.xml, localParsed.shape, {
      kind: 'solid',
      color: { kind: 'srgb', value: '654321' },
    }, PART_URI)).toBe(true);
    const localUpdated = localParsed.xml.serialize();
    expect(localUpdated).toContain(
      `<d:solidFill xmlns:d="${DRAWING_NAMESPACE}">` +
      '<d:srgbClr val="654321"/></d:solidFill><p:extLst/>',
    );
    const reparsed = parse(localUpdated);
    expect(readShapeFill(reparsed.xml, reparsed.shape)).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '654321' },
    });
  });

  it('does not patch unsafe or ambiguous shape property containers', () => {
    const sources = [
      fixture(''),
      fixture('<p:spPr/><p:spPr/>'),
      fixture('<x:spPr xmlns:x="urn:wrong"><a:noFill/></x:spPr>'),
      fixture(properties('<a:noFill/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')),
      fixture(properties('', '')),
      fixture(properties('', '<a:prstGeom prst="rect"/><a:custGeom/>')),
    ];
    for (const source of sources) {
      const { xml, shape } = parse(source);
      expect(() => replaceShapeFill(
        xml,
        shape,
        { kind: 'none' },
        PART_URI,
      ), source).toThrow(ModelParseError);
      expect(xml.changed, source).toBe(false);
      expect(xml.serialize(), source).toBe(source);
    }
  });

  it('treats clear on an absent fill as an exact no-op without requiring geometry', () => {
    const source = fixture('<p:spPr keep="EMPTY"/>');
    const { xml, shape } = parse(source);
    expect(replaceShapeFill(xml, shape, undefined, PART_URI)).toBe(false);
    expect(xml.changed).toBe(false);
    expect(xml.serialize()).toBe(source);
  });
});
