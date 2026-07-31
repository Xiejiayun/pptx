import { LosslessXmlDocument } from '@pptx/lossless-xml';
import { describe, expect, it } from 'vitest';
import { ModelParseError } from './errors.js';
import {
  readShapeLine,
  replaceShapeLine,
} from './shape-line.internal.js';

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

function properties(
  line = '<a:ln/>',
  fill = '<a:noFill/>',
  geometry = '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
): string {
  return '<p:spPr keep="PROPERTIES"><a:xfrm><a:off x="1" y="2"/>' +
    `<a:ext cx="3" cy="4"/></a:xfrm>${geometry}${fill}${line}` +
    '<a:effectLst/><a:scene3d/><a:sp3d/>' +
    '<a:extLst><a:ext uri="urn:keep"><x:opaque xmlns:x="urn:test"/></a:ext></a:extLst>' +
    '</p:spPr>';
}

function parse(source: string) {
  const xml = LosslessXmlDocument.parse(source);
  const shape = xml.roots[0];
  if (!shape) throw new Error('Fixture has no shape');
  return { xml, shape };
}

describe('shape line reader', () => {
  it('reads detached direct none, defaults, scheme alpha, width, dash, and alternate prefixes', () => {
    const cases = [
      {
        source: fixture(properties(
          '<a:ln cap="flat"><a:noFill/><a:headEnd type="triangle"/>' +
          '<a:tailEnd type="arrow"/></a:ln>',
        )),
        expected: { kind: 'none' },
      },
      {
        source: fixture(properties(
          '<a:ln><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></a:ln>',
        )),
        expected: {
          kind: 'line',
          color: { kind: 'srgb', value: 'FF0000' },
          width: 1,
          dash: 'solid',
        },
      },
      {
        source: fixture(properties(
          '<a:ln w="31750" cap="flat" cmpd="sng" algn="ctr">' +
          '<a:solidFill><a:schemeClr val="accent2"><a:alpha val="75000"/>' +
          '</a:schemeClr></a:solidFill><a:prstDash val="dashDot"/>' +
          '<a:round/><a:headEnd type="triangle"/><a:tailEnd type="arrow"/>' +
          '<a:extLst><a:ext uri="urn:line"/></a:extLst></a:ln>',
        )),
        expected: {
          kind: 'line',
          color: { kind: 'scheme', value: 'accent2' },
          transparency: 25,
          width: 2.5,
          dash: 'dashDot',
        },
      },
      {
        source:
          `<q:sp xmlns:q="${PRESENTATION_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}">` +
          '<q:spPr><d:prstGeom prst="ellipse"><d:avLst/></d:prstGeom>' +
          '<d:noFill/><d:ln w="0"><d:solidFill><d:srgbClr val="112233">' +
          '<d:alpha val="49445"/></d:srgbClr></d:solidFill>' +
          '<d:prstDash val="sysDot"/></d:ln></q:spPr></q:sp>',
        expected: {
          kind: 'line',
          color: { kind: 'srgb', value: '112233' },
          transparency: 50.555,
          width: 0,
          dash: 'sysDot',
        },
      },
    ];

    for (const { source, expected } of cases) {
      const { xml, shape } = parse(source);
      const first = readShapeLine(xml, shape);
      const second = readShapeLine(xml, shape);
      expect(first, source).toEqual(expected);
      expect(second, source).toEqual(expected);
      expect(first).not.toBe(second);
      if (first?.kind === 'line' && second?.kind === 'line') {
        expect(first.color).not.toBe(second.color);
      }
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });

  it('returns undefined for absent, empty, unsupported, nested, ambiguous, or unsafe containers', () => {
    const sources = [
      fixture(properties('')),
      fixture(properties('<a:ln/>')),
      fixture(properties('<a:ln><a:gradFill/></a:ln>')),
      fixture(properties('<a:ln><a:blipFill/></a:ln>')),
      fixture(properties('<a:ln><a:pattFill/></a:ln>')),
      fixture(properties('<a:ln><a:grpFill/></a:ln>')),
      fixture(properties(
        '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
        '<a:custDash/></a:ln>',
      )),
      fixture('<p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
        '<a:effectLst><a:ln><a:solidFill><a:srgbClr val="FF0000"/>' +
        '</a:solidFill></a:ln></a:effectLst></p:spPr>'),
      fixture('<p:spPr/><p:spPr/>'),
      fixture('<x:spPr xmlns:x="urn:wrong"><a:ln><a:noFill/></a:ln></x:spPr>'),
      fixture(properties('<a:ln><a:noFill/></a:ln><a:ln><a:noFill/></a:ln>')),
      fixture(properties('<x:ln xmlns:x="urn:wrong"><x:noFill/></x:ln>')),
      fixture(properties('<a:ln><a:noFill/></a:ln>'), { rootName: 'p:pic' }),
      fixture(properties('<a:ln><a:noFill/></a:ln>'), { presentationNamespace: 'urn:wrong' }),
    ];
    for (const source of sources) {
      const { xml, shape } = parse(source);
      expect(readShapeLine(xml, shape), source).toBeUndefined();
      expect(xml.changed, source).toBe(false);
      expect(xml.serialize(), source).toBe(source);
    }
  });

  it('returns undefined for malformed owned state and namespace lookalikes', () => {
    const lines = [
      '<a:ln w="12700"><a:noFill/></a:ln>',
      '<a:ln><a:noFill/><a:prstDash val="solid"/></a:ln>',
      '<a:ln><a:noFill/><a:solidFill><a:srgbClr val="FF0000"/>' +
        '</a:solidFill></a:ln>',
      '<a:ln><a:solidFill/></a:ln>',
      '<a:ln><a:solidFill><a:srgbClr val="FFF"/></a:solidFill></a:ln>',
      '<a:ln><a:solidFill><a:schemeClr val="unknown"/></a:solidFill></a:ln>',
      '<a:ln><a:solidFill><a:srgbClr val="FF0000"><a:tint val="50000"/>' +
        '</a:srgbClr></a:solidFill></a:ln>',
      '<a:ln w="1.5"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>',
      '<a:ln w="20116801"><a:solidFill><a:srgbClr val="FF0000"/>' +
        '</a:solidFill></a:ln>',
      '<a:ln cap="round"><a:solidFill><a:srgbClr val="FF0000"/>' +
        '</a:solidFill></a:ln>',
      '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
        '<a:prstDash val="dot"/></a:ln>',
      '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
        '<a:prstDash val="dash"/><a:prstDash val="solid"/></a:ln>',
      '<a:ln><x:solidFill xmlns:x="urn:wrong"><x:srgbClr val="FF0000"/>' +
        '</x:solidFill></a:ln>',
      `<a:ln><x:solidFill xmlns:x="${DRAWING_NAMESPACE}">` +
        '<x:srgbClr val="FF0000"/></x:solidFill></a:ln>',
      '<a:ln><a:solidFill xmlns:a="urn:wrong"><a:srgbClr val="FF0000"/>' +
        '</a:solidFill></a:ln>',
    ];
    for (const line of lines) {
      const source = fixture(properties(line));
      const { xml, shape } = parse(source);
      expect(readShapeLine(xml, shape), line).toBeUndefined();
      expect(xml.serialize()).toBe(source);
    }
  });
});

describe('shape line replacement', () => {
  it('preserves exact bytes for same-value assignments', () => {
    const sources = [
      fixture(properties(
        '<a:ln cap="flat"><a:noFill/><a:headEnd type="triangle"/></a:ln>',
      )),
      fixture(properties(
        '<a:ln w="31750"><a:solidFill><a:srgbClr val="FF0000"/>' +
        '</a:solidFill><a:prstDash val="dash"/><a:tailEnd type="arrow"/></a:ln>',
      )),
    ];
    const values = [
      { kind: 'none' } as const,
      {
        kind: 'line',
        color: { kind: 'srgb', value: 'FF0000' },
        width: 2.5,
        dash: 'dash',
      } as const,
    ];
    for (let index = 0; index < sources.length; index += 1) {
      const source = sources[index]!;
      const { xml, shape } = parse(source);
      expect(replaceShapeLine(xml, shape, values[index]!, PART_URI)).toBe(false);
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });

  it('replaces and clears owned state while preserving attributes, arrows, joins, and extensions', () => {
    const fillChoices = [
      '<a:noFill/>',
      '<a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill>',
      '<a:gradFill><a:gsLst/></a:gradFill>',
      '<a:blipFill><a:blip r:embed="rId7" xmlns:r="urn:r"/></a:blipFill>',
      '<a:pattFill prst="pct10"/>',
      '<a:grpFill/>',
    ];
    for (const choice of fillChoices) {
      const line = '<a:ln w="9" cap="flat" data-keep="LINE">' +
        `${choice}<a:custDash><a:ds d="1" sp="1"/></a:custDash>` +
        '<a:round/><a:headEnd type="triangle" w="lg" len="sm"/>' +
        '<a:tailEnd type="arrow"/><a:extLst><a:ext uri="urn:line">' +
        '<x:lineKeep xmlns:x="urn:test"/></a:ext></a:extLst></a:ln>';
      const source = fixture(properties(line));
      const replaced = parse(source);
      expect(replaceShapeLine(replaced.xml, replaced.shape, {
        kind: 'line',
        color: { kind: 'scheme', value: 'accent4' },
        transparency: 40,
        width: 2.5,
        dash: 'lgDashDot',
      }, PART_URI), choice).toBe(true);
      const updated = replaced.xml.serialize();
      expect(updated).toContain(
        '<a:ln w="31750" cap="flat" data-keep="LINE">' +
        '<a:solidFill><a:schemeClr val="accent4"><a:alpha val="60000"/>' +
        '</a:schemeClr></a:solidFill><a:prstDash val="lgDashDot"/>',
      );
      expect(updated).toContain('<a:round/><a:headEnd type="triangle" w="lg" len="sm"/>');
      expect(updated).toContain('<a:tailEnd type="arrow"/>');
      expect(updated).toContain('<x:lineKeep xmlns:x="urn:test"/>');
      expect(updated).toContain('<p:cNvPr id="7" name="Keep"/>');
      expect(updated).toContain('<a:t>KEEP</a:t>');

      const cleared = parse(source);
      expect(replaceShapeLine(cleared.xml, cleared.shape, undefined, PART_URI), choice)
        .toBe(true);
      const clearedXml = cleared.xml.serialize();
      expect(clearedXml).toContain(
        '<a:ln cap="flat" data-keep="LINE"><a:round/>' +
        '<a:headEnd type="triangle" w="lg" len="sm"/><a:tailEnd type="arrow"/>',
      );
      const clearedLine = /<a:ln\b[\s\S]*?<\/a:ln>/.exec(clearedXml)?.[0];
      expect(clearedLine).toBeDefined();
      expect(clearedLine).not.toContain(choice);
      expect(clearedLine).not.toContain('<a:custDash>');
      expect(clearedXml).toContain('<x:lineKeep xmlns:x="urn:test"/>');
    }
  });

  it('writes none into self-closing or styled lines and removes width and dash', () => {
    for (const line of [
      '<a:ln/>',
      '<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/>' +
        '</a:solidFill><a:prstDash val="dash"/><a:bevel/>' +
        '<a:headEnd type="triangle"/></a:ln>',
    ]) {
      const source = fixture(properties(line));
      const { xml, shape } = parse(source);
      expect(replaceShapeLine(xml, shape, { kind: 'none' }, PART_URI)).toBe(true);
      const updated = xml.serialize();
      expect(updated).toContain('<a:ln><a:noFill/>');
      expect(updated).not.toContain('<a:prstDash');
      expect(updated).not.toContain('<a:ln w=');
      if (line.includes('headEnd')) {
        expect(updated).toContain('<a:bevel/><a:headEnd type="triangle"/>');
      }
    }
  });

  it('inserts a missing line after shape fill and before effects with safe namespaces', () => {
    const cases = [
      {
        source: fixture(properties('', '<a:noFill/>')),
        expected:
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>' +
          '<a:ln w="31750"><a:solidFill><a:srgbClr val="123456"/>' +
          '</a:solidFill><a:prstDash val="dashDot"/></a:ln><a:effectLst/>',
      },
      {
        source: fixture(properties('', '')),
        expected:
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
          '<a:ln w="31750"><a:solidFill><a:srgbClr val="123456"/>' +
          '</a:solidFill><a:prstDash val="dashDot"/></a:ln><a:effectLst/>',
      },
      {
        source:
          `<q:sp xmlns:q="${PRESENTATION_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}">` +
          '<q:spPr><d:prstGeom prst="rect"><d:avLst/></d:prstGeom><d:noFill/>' +
          '<d:effectLst/></q:spPr></q:sp>',
        expected:
          '<d:noFill/><d:ln w="31750"><d:solidFill><d:srgbClr val="123456"/>' +
          '</d:solidFill><d:prstDash val="dashDot"/></d:ln><d:effectLst/>',
      },
      {
        source:
          `<p:sp xmlns:p="${PRESENTATION_NAMESPACE}"><p:spPr>` +
          `<d:prstGeom xmlns:d="${DRAWING_NAMESPACE}" prst="rect"><d:avLst/>` +
          '</d:prstGeom><p:extLst/></p:spPr></p:sp>',
        expected:
          `</d:prstGeom><d:ln xmlns:d="${DRAWING_NAMESPACE}" w="31750">` +
          '<d:solidFill><d:srgbClr val="123456"/></d:solidFill>' +
          '<d:prstDash val="dashDot"/></d:ln><p:extLst/>',
      },
    ];
    for (const { source, expected } of cases) {
      const { xml, shape } = parse(source);
      expect(replaceShapeLine(xml, shape, {
        kind: 'line',
        color: { kind: 'srgb', value: '123456' },
        width: 2.5,
        dash: 'dashDot',
      }, PART_URI), source).toBe(true);
      expect(xml.serialize(), source).toContain(expected);
    }
  });

  it('rejects repeated, namespace-lookalike, or unsafe containers before patching', () => {
    const sources = [
      fixture(''),
      fixture('<p:spPr/><p:spPr/>'),
      fixture('<x:spPr xmlns:x="urn:wrong"><a:ln/></x:spPr>'),
      fixture(properties('<a:ln/><a:ln/>')),
      fixture(properties('<x:ln xmlns:x="urn:wrong"/>')),
      fixture(properties(
        '<a:ln><a:noFill/><a:solidFill><a:srgbClr val="FF0000"/>' +
        '</a:solidFill></a:ln>',
      )),
      fixture(properties(
        '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
        '<a:prstDash val="dash"/><a:custDash/></a:ln>',
      )),
      fixture(properties(
        '<a:ln w="1" w="2"><a:solidFill><a:srgbClr val="FF0000"/>' +
        '</a:solidFill></a:ln>',
      )),
      fixture(properties(
        '<a:ln w="1.5"><a:solidFill><a:srgbClr val="FF0000"/>' +
        '</a:solidFill></a:ln>',
      )),
      fixture(properties('<a:ln><a:solidFill/></a:ln>')),
      fixture(properties(
        '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
        '<a:prstDash val="dot"/></a:ln>',
      )),
      fixture(properties(
        '<a:ln><x:solidFill xmlns:x="urn:wrong"><x:srgbClr val="FF0000"/>' +
        '</x:solidFill></a:ln>',
      )),
      fixture(properties(
        '<a:ln><a:solidFill><a:srgbClr xmlns:a="urn:wrong" val="FF0000"/>' +
        '</a:solidFill></a:ln>',
      )),
      fixture(properties(
        '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>' +
        '<x:prstDash xmlns:x="urn:wrong" val="dash"/></a:ln>',
      )),
      fixture(properties('', '<a:noFill/><a:solidFill><a:srgbClr val="FF0000"/>' +
        '</a:solidFill>')),
      fixture(properties('', '', '')),
      fixture(properties('', '', '<a:prstGeom prst="rect"/><a:custGeom/>')),
      fixture('<p:spPr><a:effectLst/><a:prstGeom prst="rect"/>' +
        '<a:noFill/></p:spPr>'),
    ];
    for (const source of sources) {
      const { xml, shape } = parse(source);
      expect(() => replaceShapeLine(xml, shape, {
        kind: 'none',
      }, PART_URI), source).toThrow(ModelParseError);
      expect(xml.changed, source).toBe(false);
      expect(xml.serialize(), source).toBe(source);
    }
  });

  it('treats clear on absent or already-empty owned state as an exact no-op', () => {
    for (const source of [
      fixture('<p:spPr keep="EMPTY"/>'),
      fixture(properties('<a:ln><a:headEnd type="triangle"/></a:ln>')),
    ]) {
      const { xml, shape } = parse(source);
      expect(replaceShapeLine(xml, shape, undefined, PART_URI)).toBe(false);
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });
});
