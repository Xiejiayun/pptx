import { LosslessXmlDocument } from '@pptx/lossless-xml';
import type { Relationship } from '@pptx/opc';
import { describe, expect, it } from 'vitest';
import { ModelParseError } from './errors.js';
import {
  normalizeHyperlink,
  readShapeHyperlink,
  readTextRunHyperlink,
  relationshipReferenceCount,
  removeDrawingHyperlinkReferences,
  requireShapeHyperlinkRelationshipId,
  renderShapeHyperlink,
  replaceShapeHyperlinkElement,
  replaceTextRunHyperlinkElement,
  shapeHyperlinksEqual,
  type ShapeHyperlinkReadContext,
} from './shape-hyperlink.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const HYPERLINK_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}/hyperlink`;
const SLIDE_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}/slide`;
const PART_URI = '/ppt/slides/slide1.xml';

function fixture(
  hyperlinkChildren = '',
  options: {
    readonly cNvPr?: string;
    readonly root?: string;
  } = {},
): string {
  if (options.root) return options.root;
  const nonVisual = options.cNvPr ??
    `<p:cNvPr id="7" name="Keep" descr="Description">${hyperlinkChildren}</p:cNvPr>`;
  return `<p:sp xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" ` +
    `xmlns:r="${RELATIONSHIP_NAMESPACE}" xmlns:x="urn:test">` +
    `<p:nvSpPr keep="NV">${nonVisual}<p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    '<p:spPr keep="PROPERTIES"><a:xfrm><a:off x="1" y="2"/>' +
    '<a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/>' +
    '</a:prstGeom><a:noFill/><a:ln><a:headEnd type="triangle"/></a:ln>' +
    '<a:effectLst/><a:extLst><a:ext uri="urn:shape"><x:keep/></a:ext></a:extLst>' +
    '</p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>KEEP</a:t></a:r></a:p>' +
    '</p:txBody></p:sp>';
}

function parse(source: string) {
  const xml = LosslessXmlDocument.parse(source);
  const shape = xml.roots[0];
  if (!shape) throw new Error('Fixture has no shape');
  return { xml, shape };
}

function external(
  target = 'https://example.com?a=1&b=2',
  id = 'rId7',
): Relationship {
  return {
    id,
    type: HYPERLINK_RELATIONSHIP,
    target,
    targetMode: 'External',
  };
}

function internal(
  resolvedTarget = '/ppt/slides/slide2.xml',
  id = 'rId7',
): Relationship {
  return {
    id,
    type: SLIDE_RELATIONSHIP,
    target: resolvedTarget.slice(resolvedTarget.lastIndexOf('/') + 1),
    targetMode: 'Internal',
    resolvedTarget,
  };
}

function context(
  relationships: readonly Relationship[],
  slidePartUris = ['/ppt/slides/slide1.xml', '/ppt/slides/slide2.xml'],
): ShapeHyperlinkReadContext {
  return { relationships, slidePartUris };
}

describe('shape hyperlink normalization', () => {
  it('normalizes, freezes, and detaches URL and slide targets with direct tooltip state', () => {
    const urlInput: { url: string; tooltip?: string } = {
      url: 'https://example.com?a=1&b=2',
      tooltip: 'Visit & learn',
    };
    const url = normalizeHyperlink(urlInput, 'Shape hyperlink');
    urlInput.url = 'https://changed.example';
    delete urlInput.tooltip;
    expect(url).toEqual({
      url: 'https://example.com?a=1&b=2',
      tooltip: 'Visit & learn',
    });
    expect(Object.isFrozen(url)).toBe(true);

    expect(normalizeHyperlink({ slide: 2, tooltip: '' }, 'Shape hyperlink'))
      .toEqual({ slide: 2, tooltip: '' });
    expect(normalizeHyperlink({ url: undefined, slide: 1 }, 'Shape hyperlink'))
      .toEqual({ slide: 1 });

    const nullPrototype = Object.create(null) as Record<string, unknown>;
    nullPrototype.url = 'mailto:test@example.com';
    expect(normalizeHyperlink(nullPrototype, 'Shape hyperlink'))
      .toEqual({ url: 'mailto:test@example.com' });
  });

  it('rejects unsafe objects, ambiguous targets, coercions, aliases, and invalid XML strings', () => {
    class HyperlinkOptions {
      url = 'https://example.com';
    }
    const inherited = Object.create({ url: 'https://example.com' }) as Record<string, unknown>;
    const symbol = { url: 'https://example.com', [Symbol('target')]: 1 };
    const invalidValues = [
      undefined,
      null,
      true,
      1,
      'https://example.com',
      [],
      new Date(),
      new HyperlinkOptions(),
      inherited,
      symbol,
      {},
      { url: undefined, slide: undefined },
      { url: 'https://example.com', slide: 2 },
      { url: '' },
      { url: 42 },
      { url: 'https://example.com\u0000' },
      { slide: 0 },
      { slide: -1 },
      { slide: 1.5 },
      { slide: Number.NaN },
      { slide: Number.POSITIVE_INFINITY },
      { slide: Number.MAX_SAFE_INTEGER + 1 },
      { slide: '2' },
      { slide: true },
      { slide: 2, tooltip: 1 },
      { slide: 2, tooltip: '\u0001' },
      { url: 'https://example.com', unknown: true },
      { url: 'https://example.com', _rId: 7 },
      { target: 'https://example.com' },
      { kind: 'url', url: 'https://example.com' },
    ];
    for (const value of invalidValues) {
      expect(() => normalizeHyperlink(value, 'Shape hyperlink'), String(value))
        .toThrow(TypeError);
    }

    let invoked = false;
    const accessor = Object.defineProperty({}, 'url', {
      enumerable: true,
      get() {
        invoked = true;
        return 'https://example.com';
      },
    });
    expect(() => normalizeHyperlink(accessor, 'Shape hyperlink')).toThrow(TypeError);
    expect(invoked).toBe(false);
  });

  it('renders escaped canonical click elements and compares direct tooltip presence', () => {
    expect(renderShapeHyperlink(
      { url: 'https://example.com?a=1&b=2', tooltip: '' },
      'rId7',
      { drawing: 'a', relationship: 'r' },
    )).toBe('<a:hlinkClick r:id="rId7" tooltip=""/>');
    expect(renderShapeHyperlink(
      { slide: 2, tooltip: 'Next & <details> "now"' },
      'rId8',
      { drawing: 'd', relationship: 'rel' },
    )).toBe(
      '<d:hlinkClick rel:id="rId8" tooltip="Next &amp; &lt;details&gt; &quot;now&quot;" ' +
      'action="ppaction://hlinksldjump"/>',
    );
    expect(shapeHyperlinksEqual(undefined, undefined)).toBe(true);
    expect(shapeHyperlinksEqual({ url: 'https://example.com' }, {
      url: 'https://example.com',
    })).toBe(true);
    expect(shapeHyperlinksEqual({ url: 'https://example.com' }, {
      url: 'https://example.com',
      tooltip: '',
    })).toBe(false);
    expect(shapeHyperlinksEqual({ slide: 2, tooltip: '' }, {
      slide: 2,
      tooltip: '',
    })).toBe(true);
    expect(shapeHyperlinksEqual({ slide: 2 }, { url: '2' })).toBe(false);
  });
});

describe('shape hyperlink reader', () => {
  it('requires a structurally safe container and exposes the direct relationship ID', () => {
    const linked = parse(fixture('<a:hlinkClick r:id="rId7"/>'));
    expect(requireShapeHyperlinkRelationshipId(linked.shape, PART_URI)).toBe('rId7');

    const absent = parse(fixture());
    expect(requireShapeHyperlinkRelationshipId(absent.shape, PART_URI)).toBeUndefined();

    for (const source of [
      fixture('<a:hlinkClick x:id="rId7"/>'),
      fixture('<a:hlinkClick r:id="rId7"/><a:hlinkClick r:id="rId8"/>'),
      fixture('', { cNvPr: '<p:cNvPr id="7"/><p:cNvPr id="8"/>' }),
    ]) {
      const malformed = parse(source);
      expect(() => requireShapeHyperlinkRelationshipId(malformed.shape, PART_URI))
        .toThrow(ModelParseError);
    }
  });

  it('reads detached external URL snapshots with absent, empty, and custom tooltips', () => {
    const cases = [
      {
        click: '<a:hlinkClick r:id="rId7"/>',
        expected: { url: 'https://example.com?a=1&b=2' },
      },
      {
        click: '<a:hlinkClick r:id="rId7" tooltip=""/>',
        expected: { url: 'https://example.com?a=1&b=2', tooltip: '' },
      },
      {
        click: '<a:hlinkClick r:id="rId7" tooltip="Visit &amp; learn" history="0">' +
          '<a:snd r:embed="rIdSound"/><a:extLst><a:ext uri="urn:keep"/></a:extLst>' +
          '</a:hlinkClick>',
        expected: {
          url: 'https://example.com?a=1&b=2',
          tooltip: 'Visit & learn',
        },
      },
    ];
    for (const { click, expected } of cases) {
      const source = fixture(click);
      const { xml, shape } = parse(source);
      const first = readShapeHyperlink(xml, shape, context([external()]));
      const second = readShapeHyperlink(xml, shape, context([external()]));
      expect(first).toEqual(expected);
      expect(second).toEqual(expected);
      expect(first).not.toBe(second);
      expect(Object.isFrozen(first)).toBe(true);
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });

  it('reads current one-based internal and self targets with alternate prefixes', () => {
    const source =
      `<q:sp xmlns:q="${PRESENTATION_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}" ` +
      `xmlns:rel="${RELATIONSHIP_NAMESPACE}"><q:nvSpPr><q:cNvPr id="7" name="Keep">` +
      '<d:hlinkClick rel:id="link9" tooltip="Next" ' +
      'action="ppaction://hlinksldjump"/></q:cNvPr><q:cNvSpPr/><q:nvPr/>' +
      '</q:nvSpPr><q:spPr><d:prstGeom prst="rect"><d:avLst/></d:prstGeom>' +
      '</q:spPr></q:sp>';
    const { xml, shape } = parse(source);
    expect(readShapeHyperlink(
      xml,
      shape,
      context([internal('/ppt/slides/custom.xml', 'link9')], [
        '/ppt/slides/slide1.xml',
        '/ppt/slides/other.xml',
        '/ppt/slides/custom.xml',
      ]),
    )).toEqual({ slide: 3, tooltip: 'Next' });

    const selfSource = fixture(
      '<a:hlinkClick r:id="rId7" action="ppaction://hlinksldjump"/>',
    );
    const self = parse(selfSource);
    expect(readShapeHyperlink(
      self.xml,
      self.shape,
      context([internal('/ppt/slides/slide1.xml')]),
    )).toEqual({ slide: 1 });
  });

  it('returns undefined for absent, malformed, ambiguous, or unsupported state', () => {
    const baseContext = context([external()]);
    const cases: readonly {
      readonly source: string;
      readonly readContext?: ShapeHyperlinkReadContext;
    }[] = [
      { source: fixture() },
      {
        source: fixture(
          '<a:hlinkClick r:id="rId7"/><a:hlinkClick r:id="rId8"/>',
        ),
      },
      { source: fixture('<x:hlinkClick r:id="rId7"/>') },
      { source: fixture('<a:hlinkClick/>') },
      {
        source: fixture(
          '<a:hlinkClick r:id="rId7" rel:id="rId7" ' +
          `xmlns:rel="${RELATIONSHIP_NAMESPACE}"/>`,
        ),
      },
      { source: fixture('<a:hlinkClick r:id="rId7" x:tooltip="Wrong"/>') },
      {
        source: fixture(
          '<a:hlinkClick r:id="rId7" tooltip="One" tooltip="Two"/>',
        ),
      },
      { source: fixture('<a:hlinkClick r:id="rId7" x:action="Wrong"/>') },
      {
        source: fixture(
          '<a:hlinkClick r:id="rId7"><a:hlinkClick r:id="rId8"/></a:hlinkClick>',
        ),
      },
      {
        source: fixture(
          '<a:hlinkClick r:id="rId7"><a:extLst/><a:snd r:embed="rIdSound"/>' +
          '</a:hlinkClick>',
        ),
      },
      {
        source: fixture(
          '<a:hlinkHover r:id="rId8"/><a:hlinkClick r:id="rId7"/>',
        ),
      },
      {
        source: fixture('<a:hlinkClick r:id="rId7"/>'),
        readContext: context([external(), external('https://two.example', 'rId7')]),
      },
      {
        source: fixture('<a:hlinkClick r:id="rId7"/>'),
        readContext: context([{ ...external(), target: '' }]),
      },
      {
        source: fixture('<a:hlinkClick r:id="rId7"/>'),
        readContext: context([{ ...external(), type: SLIDE_RELATIONSHIP }]),
      },
      {
        source: fixture('<a:hlinkClick r:id="rId7"/>'),
        readContext: context([{ ...external(), targetMode: 'Internal' }]),
      },
      {
        source: fixture(
          '<a:hlinkClick r:id="rId7" action="ppaction://hlinksldjump"/>',
        ),
        readContext: context([external()]),
      },
      {
        source: fixture('<a:hlinkClick r:id="rId7"/>'),
        readContext: context([internal()]),
      },
      {
        source: fixture(
          '<a:hlinkClick r:id="rId7" action="ppaction://hlinksldjump"/>',
        ),
        readContext: context([internal('/ppt/slides/missing.xml')]),
      },
      {
        source: fixture(
          '<a:hlinkClick r:id="rId7" action="ppaction://hlinksldjump"/>',
        ),
        readContext: context([internal()], [
          '/ppt/slides/slide1.xml',
          '/ppt/slides/slide2.xml',
          '/ppt/slides/slide2.xml',
        ]),
      },
      {
        source: fixture('', {
          root: `<x:sp xmlns:x="urn:test"><x:nvSpPr><x:cNvPr>` +
            '<x:hlinkClick id="rId7"/></x:cNvPr></x:nvSpPr></x:sp>',
        }),
      },
      {
        source: fixture('', {
          cNvPr: '<p:cNvPr id="7"/><p:cNvPr id="8"/>',
        }),
      },
    ];

    for (const { source, readContext = baseContext } of cases) {
      const { xml, shape } = parse(source);
      expect(readShapeHyperlink(xml, shape, readContext), source).toBeUndefined();
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });
});

describe('text run hyperlink reader', () => {
  function properties(click: string, rootName = 'a:rPr'): ReturnType<typeof parse> {
    return parse(
      `<${rootName} xmlns:a="${DRAWING_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}" ` +
      `xmlns:x="urn:test">${click}</${rootName}>`,
    );
  }

  it('reads PptxGenJS external and canonical internal run links', () => {
    const url = properties(
      '<a:hlinkClick r:id="rId7" tooltip="" action="" invalidUrl="" ' +
      'tgtFrame="" history="1" highlightClick="0" endSnd="0">' +
      '<a:extLst><a:ext uri="urn:pptxgenjs"><x:keep/></a:ext></a:extLst>' +
      '</a:hlinkClick>',
    );
    expect(readTextRunHyperlink(url.shape, context([external()])))
      .toEqual({ url: 'https://example.com?a=1&b=2', tooltip: '' });

    const slide = properties(
      '<a:hlinkClick r:id="rId7" tooltip="Next" ' +
      'action="ppaction://hlinksldjump"/>',
    );
    expect(readTextRunHyperlink(slide.shape, context([internal()])))
      .toEqual({ slide: 2, tooltip: 'Next' });
    expect(url.xml.changed).toBe(false);
    expect(slide.xml.changed).toBe(false);
  });

  it('returns undefined for malformed, ambiguous, or unsupported run state', () => {
    const cases: readonly {
      readonly click: string;
      readonly readContext?: ShapeHyperlinkReadContext;
      readonly rootName?: string;
    }[] = [
      { click: '' },
      { click: '<x:hlinkClick r:id="rId7"/>' },
      { click: '<a:hlinkClick x:id="rId7"/>' },
      { click: '<a:hlinkClick r:id="rId7"/><a:hlinkClick r:id="rId8"/>' },
      { click: '<a:hlinkClick r:id="rId7" action="ppaction://unsupported"/>' },
      { click: '<a:hlinkClick r:id="rId7"><a:unknown/></a:hlinkClick>' },
      {
        click: '<a:hlinkClick r:id="rId7"/>',
        readContext: context([external(), external('https://duplicate.example', 'rId7')]),
      },
      {
        click: '<a:hlinkClick r:id="rId7"/>',
        readContext: context([{ ...external(), targetMode: 'Internal' }]),
      },
      {
        click: '<a:hlinkClick r:id="rId7"/>',
        readContext: context([{ ...external(), target: '' }]),
      },
      {
        click: '<a:hlinkClick r:id="rId7" action="ppaction://hlinksldjump"/>',
        readContext: context([internal('/ppt/slides/missing.xml')]),
      },
      { click: '<a:hlinkClick r:id="rId7"/>', rootName: 'x:rPr' },
    ];
    for (const { click, readContext = context([external()]), rootName } of cases) {
      const parsed = properties(click, rootName);
      expect(readTextRunHyperlink(parsed.shape, readContext), click).toBeUndefined();
      expect(parsed.xml.changed).toBe(false);
    }
  });
});

describe('shape hyperlink element replacement', () => {
  it('inserts canonical click elements into expanded and self-closing non-visual properties', () => {
    const expandedSource = fixture(
      '<a:hlinkHover r:id="rIdHover"/><a:extLst><a:ext uri="urn:keep"/></a:extLst>',
    );
    const expanded = parse(expandedSource);
    expect(replaceShapeHyperlinkElement(
      expanded.xml,
      expanded.shape,
      { url: 'https://example.com', tooltip: 'Visit' },
      'rId7',
      PART_URI,
    )).toBe(true);
    expect(expanded.xml.serialize()).toContain(
      '<a:hlinkClick r:id="rId7" tooltip="Visit"/>' +
      '<a:hlinkHover r:id="rIdHover"/>',
    );
    expect(expanded.xml.serialize()).toContain('<a:headEnd type="triangle"/>');

    const selfClosingSource = fixture('', {
      cNvPr: '<p:cNvPr id="7" name="Self closing"/>',
    });
    const selfClosing = parse(selfClosingSource);
    expect(replaceShapeHyperlinkElement(
      selfClosing.xml,
      selfClosing.shape,
      { slide: 2 },
      'rId8',
      PART_URI,
    )).toBe(true);
    expect(selfClosing.xml.serialize()).toContain(
      '<p:cNvPr id="7" name="Self closing"><a:hlinkClick r:id="rId8" ' +
      'action="ppaction://hlinksldjump"/></p:cNvPr>',
    );
  });

  it('patches only owned attributes while preserving extra click, hover, and shape bytes', () => {
    const click = '<a:hlinkClick r:id="rId7" history="0" highlightClick="1" ' +
      'tooltip="Old" action="ppaction://hlinksldjump"><a:snd r:embed="rIdSound"/>' +
      '<a:extLst><a:ext uri="urn:click"><x:keep/></a:ext></a:extLst>' +
      '</a:hlinkClick><a:hlinkHover r:id="rIdHover" tooltip="Hover"/>' +
      '<a:extLst><a:ext uri="urn:nv"><x:keep/></a:ext></a:extLst>';
    const source = fixture(click);
    const { xml, shape } = parse(source);
    expect(replaceShapeHyperlinkElement(
      xml,
      shape,
      { url: 'https://replacement.example', tooltip: '' },
      'rId9',
      PART_URI,
    )).toBe(true);
    const output = xml.serialize();
    expect(output).toContain(
      '<a:hlinkClick r:id="rId9" history="0" highlightClick="1" tooltip="">' +
      '<a:snd r:embed="rIdSound"/><a:extLst><a:ext uri="urn:click">' +
      '<x:keep/></a:ext></a:extLst></a:hlinkClick>',
    );
    expect(output).toContain('<a:hlinkHover r:id="rIdHover" tooltip="Hover"/>');
    expect(output).toContain('<a:headEnd type="triangle"/>');
    expect(output).toContain('<a:t>KEEP</a:t>');

    const cleared = parse(output);
    expect(replaceShapeHyperlinkElement(
      cleared.xml,
      cleared.shape,
      undefined,
      undefined,
      PART_URI,
    )).toBe(true);
    expect(cleared.xml.serialize()).not.toContain('<a:hlinkClick');
    expect(cleared.xml.serialize()).toContain('<a:hlinkHover r:id="rIdHover"');
    expect(cleared.xml.serialize()).not.toContain('<a:ext uri="urn:click">');
    expect(cleared.xml.serialize()).toContain('<a:ext uri="urn:nv">');
  });

  it('distinguishes direct tooltip absence/empty and performs exact no-ops', () => {
    const source = fixture('<a:hlinkClick r:id="rId7" tooltip=""/>');
    const same = parse(source);
    expect(replaceShapeHyperlinkElement(
      same.xml,
      same.shape,
      { url: 'https://example.com', tooltip: '' },
      'rId7',
      PART_URI,
    )).toBe(false);
    expect(same.xml.changed).toBe(false);
    expect(same.xml.serialize()).toBe(source);

    const removeTooltip = parse(source);
    expect(replaceShapeHyperlinkElement(
      removeTooltip.xml,
      removeTooltip.shape,
      { url: 'https://example.com' },
      'rId7',
      PART_URI,
    )).toBe(true);
    expect(removeTooltip.xml.serialize()).toContain('<a:hlinkClick r:id="rId7"/>');

    const absent = parse(fixture());
    expect(replaceShapeHyperlinkElement(
      absent.xml,
      absent.shape,
      undefined,
      undefined,
      PART_URI,
    )).toBe(false);
    expect(absent.xml.changed).toBe(false);
  });

  it('uses alternate in-scope prefixes and creates local declarations when absent', () => {
    const alternateSource =
      `<q:sp xmlns:q="${PRESENTATION_NAMESPACE}" xmlns:d="${DRAWING_NAMESPACE}" ` +
      `xmlns:rel="${RELATIONSHIP_NAMESPACE}"><q:nvSpPr><q:cNvPr id="7"/>` +
      '<q:cNvSpPr/><q:nvPr/></q:nvSpPr><q:spPr><d:prstGeom prst="rect">' +
      '<d:avLst/></d:prstGeom></q:spPr></q:sp>';
    const alternate = parse(alternateSource);
    replaceShapeHyperlinkElement(
      alternate.xml,
      alternate.shape,
      { url: 'https://example.com' },
      'rId7',
      PART_URI,
    );
    expect(alternate.xml.serialize()).toContain('<d:hlinkClick rel:id="rId7"/>');

    const missingSource =
      `<p:sp xmlns:p="${PRESENTATION_NAMESPACE}"><p:nvSpPr>` +
      '<p:cNvPr id="7"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>' +
      '</p:sp>';
    const missing = parse(missingSource);
    replaceShapeHyperlinkElement(
      missing.xml,
      missing.shape,
      { slide: 1 },
      'rId8',
      PART_URI,
    );
    expect(missing.xml.serialize()).toContain(
      `<a:hlinkClick xmlns:a="${DRAWING_NAMESPACE}" ` +
      `xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rId8" ` +
      'action="ppaction://hlinksldjump"/>',
    );
  });

  it('rejects malformed containers and unsupported click state without mutation', () => {
    const sources = [
      fixture('<a:hlinkClick r:id="rId7"/><a:hlinkClick r:id="rId8"/>'),
      fixture('<x:hlinkClick r:id="rId7"/>'),
      fixture('<a:hlinkClick r:id="rId7" action="ppaction://unsupported"/>'),
      fixture('<a:hlinkClick r:id="rId7" x:tooltip="Wrong"/>'),
      fixture('<a:hlinkClick r:id="rId7"><a:unknown/></a:hlinkClick>'),
      fixture('<a:hlinkHover r:id="rId8"/><a:hlinkClick r:id="rId7"/>'),
      fixture('', { cNvPr: '<p:cNvPr id="7"/><p:cNvPr id="8"/>' }),
    ];
    for (const source of sources) {
      const { xml, shape } = parse(source);
      expect(() => replaceShapeHyperlinkElement(
        xml,
        shape,
        { url: 'https://replacement.example' },
        'rId9',
        PART_URI,
      )).toThrow(ModelParseError);
      expect(xml.changed).toBe(false);
      expect(xml.serialize()).toBe(source);
    }
  });
});

describe('text-run hyperlink element replacement', () => {
  const runProperties = (contents = '', attributes = 'lang="en-US"') => {
    const source = `<a:rPr xmlns:a="${DRAWING_NAMESPACE}" xmlns:x="urn:test" ` +
      `${attributes}>${contents}</a:rPr>`;
    const xml = LosslessXmlDocument.parse(source);
    return { source, xml, properties: xml.roots[0]! };
  };

  it('inserts direct clicks with local relationship namespace and default underline', () => {
    const parsed = runProperties(
      '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>' +
      '<a:extLst><a:ext uri="urn:keep"><x:keep/></a:ext></a:extLst>',
    );
    expect(replaceTextRunHyperlinkElement(
      parsed.xml,
      parsed.properties,
      { url: 'https://example.com', tooltip: '' },
      'rId7',
      PART_URI,
    )).toBe(true);
    const output = parsed.xml.serialize();
    expect(output).toContain('<a:rPr xmlns:a=');
    expect(output).toContain('lang="en-US" u="sng"');
    expect(output).toContain(
      `<a:hlinkClick xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rId7" tooltip=""/>` +
      '<a:extLst>',
    );
    expect(output).toContain('<x:keep/>');

    const explicit = runProperties('', 'lang="en-US" u="none"');
    replaceTextRunHyperlinkElement(
      explicit.xml,
      explicit.properties,
      { slide: 2 },
      'rId8',
      PART_URI,
    );
    expect(explicit.xml.serialize()).toContain('u="none"');
    expect(explicit.xml.serialize()).toContain(
      'r:id="rId8" action="ppaction://hlinksldjump"',
    );
  });

  it('patches owned click state and preserves compatible imported extras', () => {
    const parsed = runProperties(
      '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>' +
      '<a:hlinkClick xmlns:r="' + RELATIONSHIP_NAMESPACE + '" r:id="rId7" ' +
      'invalidUrl="" action="" tgtFrame="" tooltip="Old" history="1" ' +
      'highlightClick="0" endSnd="0"><a:snd r:embed="rIdSound"/>' +
      '<a:extLst><a:ext uri="urn:click"><x:keep/></a:ext></a:extLst>' +
      '</a:hlinkClick>',
      'lang="en-US" u="sng"',
    );
    expect(replaceTextRunHyperlinkElement(
      parsed.xml,
      parsed.properties,
      { slide: 2, tooltip: '' },
      'rId9',
      PART_URI,
    )).toBe(true);
    const output = parsed.xml.serialize();
    expect(output).toContain('r:id="rId9"');
    expect(output).toContain('tooltip=""');
    expect(output).toContain('action="ppaction://hlinksldjump"');
    expect(output).toContain('invalidUrl=""');
    expect(output).toContain('history="1"');
    expect(output).toContain('<a:snd r:embed="rIdSound"/>');
    expect(output).toContain('<x:keep/>');

    const cleared = LosslessXmlDocument.parse(output);
    expect(replaceTextRunHyperlinkElement(
      cleared,
      cleared.roots[0]!,
      undefined,
      undefined,
      PART_URI,
    )).toBe(true);
    expect(cleared.serialize()).not.toContain('hlinkClick');
    expect(cleared.serialize()).toContain('u="sng"');
    expect(cleared.serialize()).toContain('<a:solidFill>');
  });

  it('performs exact no-ops and rejects unsafe direct run state', () => {
    const same = runProperties(
      `<a:hlinkClick xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rId7" tooltip=""/>`,
      'lang="en-US" u="sng"',
    );
    expect(replaceTextRunHyperlinkElement(
      same.xml,
      same.properties,
      { url: 'https://example.com', tooltip: '' },
      'rId7',
      PART_URI,
    )).toBe(false);
    expect(same.xml.changed).toBe(false);
    expect(same.xml.serialize()).toBe(same.source);

    const absent = runProperties();
    expect(replaceTextRunHyperlinkElement(
      absent.xml,
      absent.properties,
      undefined,
      undefined,
      PART_URI,
    )).toBe(false);
    expect(absent.xml.changed).toBe(false);

    for (const parsed of [
      runProperties(
        `<a:hlinkClick xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rId7"/>` +
        `<a:hlinkClick xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rId8"/>`,
      ),
      runProperties('<x:hlinkClick/>'),
      runProperties(
        `<a:hlinkClick xmlns:r="${RELATIONSHIP_NAMESPACE}" r:id="rId7" ` +
        'action="ppaction://unsupported"/>',
      ),
      runProperties('', 'u="sng" u="none"'),
    ]) {
      expect(() => replaceTextRunHyperlinkElement(
        parsed.xml,
        parsed.properties,
        { url: 'https://replacement.example' },
        'rId9',
        PART_URI,
      )).toThrow(ModelParseError);
      expect(parsed.xml.changed).toBe(false);
    }

    const mismatched = runProperties();
    expect(() => replaceTextRunHyperlinkElement(
      mismatched.xml,
      mismatched.properties,
      { url: 'https://example.com' },
      undefined,
      PART_URI,
    )).toThrow(TypeError);
  });
});

describe('shape hyperlink relationship references', () => {
  it('counts only expanded-name-correct relationship ID attributes', () => {
    const source =
      `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" ` +
      `xmlns:r="${RELATIONSHIP_NAMESPACE}" xmlns:rel="${RELATIONSHIP_NAMESPACE}" ` +
      'xmlns:x="urn:test"><a:hlinkClick r:id="rId7"/><a:rPr rel:id="rId7"/>' +
      '<p:cNvPr r:id="rId8"/><x:opaque x:id="rId7" id="rId7"/>' +
      '<a:hlinkHover r:id="rId7"/></p:sld>';
    const xml = LosslessXmlDocument.parse(source);
    expect(relationshipReferenceCount(xml, 'rId7')).toBe(3);
    expect(relationshipReferenceCount(xml, 'rId8')).toBe(1);
    expect(relationshipReferenceCount(xml, 'missing')).toBe(0);
    expect(xml.changed).toBe(false);
  });

  it('removes only target DrawingML click/hover elements and preserves all neighbors', () => {
    const source =
      `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" ` +
      `xmlns:r="${RELATIONSHIP_NAMESPACE}" xmlns:x="urn:test">` +
      '<p:cSld><a:hlinkClick r:id="rIdTarget" tooltip="Click"><a:snd r:embed="snd"/>' +
      '</a:hlinkClick><a:hlinkHover r:id="rIdTarget"/><a:hlinkClick r:id="rIdKeep"/>' +
      '<x:hlinkClick r:id="rIdTarget"/><a:rPr r:id="rIdTarget"/>' +
      '<x:opaque x:id="rIdTarget">KEEP</x:opaque></p:cSld></p:sld>';
    const xml = LosslessXmlDocument.parse(source);
    expect(removeDrawingHyperlinkReferences(xml, new Set(['rIdTarget']))).toBe(true);
    const output = xml.serialize();
    expect(output).not.toContain('<a:hlinkClick r:id="rIdTarget"');
    expect(output).not.toContain('<a:hlinkHover r:id="rIdTarget"');
    expect(output).toContain('<a:hlinkClick r:id="rIdKeep"/>');
    expect(output).toContain('<x:hlinkClick r:id="rIdTarget"/>');
    expect(output).toContain('<a:rPr r:id="rIdTarget"/>');
    expect(output).toContain('<x:opaque x:id="rIdTarget">KEEP</x:opaque>');

    const unchanged = LosslessXmlDocument.parse(source);
    expect(removeDrawingHyperlinkReferences(unchanged, new Set(['missing']))).toBe(false);
    expect(unchanged.changed).toBe(false);
    expect(unchanged.serialize()).toBe(source);
  });
});
