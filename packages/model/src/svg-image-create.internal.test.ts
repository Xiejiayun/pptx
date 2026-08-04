import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import type { AddSvgImageOptions, SvgImageContentType } from './image.js';
import {
  normalizeEmbeddedSvgImage,
  renderEmbeddedSvgImageXml,
} from './svg-image-create.internal.js';
import { degrees, inches } from './units.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SVG_NAMESPACE =
  'http://schemas.microsoft.com/office/drawing/2016/SVG/main';
const SVG_EXTENSION_URI = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';

describe('embedded SVG image normalization', () => {
  it('publishes detached SVG and fallback payloads with one-inch defaults', () => {
    const contentType: SvgImageContentType = 'image/svg+xml';
    const svgBytes = new Uint8Array([60, 115, 118, 103, 47, 62]);
    const fallbackPngBytes = new Uint8Array([137, 80, 78, 71]);
    const options: AddSvgImageOptions = {};
    const normalized = normalizeEmbeddedSvgImage(svgBytes, fallbackPngBytes, options);

    expect(normalized).toEqual({
      svgBytes: new Uint8Array([60, 115, 118, 103, 47, 62]),
      fallbackPngBytes: new Uint8Array([137, 80, 78, 71]),
      name: undefined,
      altText: 'preencoded.svg',
      x: 0,
      y: 0,
      width: 914_400,
      height: 914_400,
      rotation: 0,
      flipHorizontal: false,
      flipVertical: false,
    });
    expect(normalized.svgBytes).not.toBe(svgBytes);
    expect(normalized.fallbackPngBytes).not.toBe(fallbackPngBytes);
    expect(Object.isFrozen(normalized)).toBe(true);
    expect(contentType).toBe('image/svg+xml');

    if (false) {
      // @ts-expect-error SVG options do not accept a caller-selected content type
      const selectedContentType: AddSvgImageOptions = { contentType: 'image/svg+xml' };
      // @ts-expect-error SVG content type is canonical
      const invalidContentType: SvgImageContentType = 'image/svg';
      void [selectedContentType, invalidContentType];
    }

    svgBytes.fill(0);
    fallbackPngBytes.fill(0);
    expect(normalized.svgBytes).toEqual(new Uint8Array([60, 115, 118, 103, 47, 62]));
    expect(normalized.fallbackPngBytes).toEqual(new Uint8Array([137, 80, 78, 71]));
  });

  it('accepts null-prototype options and preserves explicit appearance state', () => {
    const options = Object.create(null) as Record<string, unknown>;
    options.name = '';
    options.altText = '';
    options.x = -12;
    options.y = 34;
    options.width = 56;
    options.height = 78;
    options.rotation = -2_700_000;
    options.flipHorizontal = true;
    options.flipVertical = true;
    const sourceRectangle = { left: 25.0004, top: -10.0004, right: 5, bottom: 0 };
    options.sourceRectangle = sourceRectangle;

    const normalized = normalizeEmbeddedSvgImage(
      new Uint8Array([1]),
      new Uint8Array([2]),
      options,
    );
    options.name = 'changed';
    sourceRectangle.left = 99;

    expect(normalized).toMatchObject({
      name: '',
      altText: '',
      x: -12,
      y: 34,
      width: 56,
      height: 78,
      rotation: -2_700_000,
      flipHorizontal: true,
      flipVertical: true,
      sourceRectangle: { left: 25, top: -10, right: 5, bottom: 0 },
    });
    expect(Object.isFrozen(normalized.sourceRectangle)).toBe(true);
  });

  it('accepts public transform converter output', () => {
    expect(normalizeEmbeddedSvgImage(
      new Uint8Array([1]),
      new Uint8Array([2]),
      {
        x: inches(1),
        y: inches(2),
        width: inches(3),
        height: inches(4),
        rotation: degrees(45),
        flipHorizontal: true,
      },
    )).toMatchObject({
      x: 914_400,
      y: 1_828_800,
      width: 2_743_200,
      height: 3_657_600,
      rotation: 2_700_000,
      flipHorizontal: true,
      flipVertical: false,
    });
  });

  it('resolves SVG image percentage coordinates through the shared appearance path', () => {
    const slideSize = Object.freeze({ width: inches(10), height: inches(8) });
    expect(normalizeEmbeddedSvgImage(
      new Uint8Array([1]),
      new Uint8Array([2]),
      {
        x: '12.5%',
        y: '25%',
        width: '37.5%',
        height: '50%',
      },
      slideSize,
    )).toMatchObject({
      x: inches(1.25),
      y: inches(2),
      width: inches(3.75),
      height: inches(4),
    });
  });

  it('rejects non-byte and empty SVG or fallback payloads', () => {
    const valid = new Uint8Array([1]);
    for (const value of [undefined, null, [], [1], new ArrayBuffer(1), 'image', 1]) {
      expect(() => normalizeEmbeddedSvgImage(value, valid, {}), `svg ${String(value)}`)
        .toThrow(TypeError);
      expect(() => normalizeEmbeddedSvgImage(valid, value, {}), `fallback ${String(value)}`)
        .toThrow(TypeError);
    }
    expect(() => normalizeEmbeddedSvgImage(new Uint8Array(), valid, {})).toThrow(RangeError);
    expect(() => normalizeEmbeddedSvgImage(valid, new Uint8Array(), {})).toThrow(RangeError);
  });

  it('rejects unsafe option containers and descriptors without reading accessors', () => {
    class Options {
      name = 'unsafe';
    }
    const inherited = Object.create({ name: 'unsafe' });
    const symbol = { [Symbol('unsafe')]: true };
    const unknown = { contentType: 'image/svg+xml' };
    let reads = 0;
    const accessor = Object.defineProperty({}, 'name', {
      get: () => {
        reads += 1;
        return 'unsafe';
      },
    });

    for (const options of [null, [], new Options(), inherited, symbol, unknown, accessor]) {
      expect(() => normalizeEmbeddedSvgImage(
        new Uint8Array([1]),
        new Uint8Array([2]),
        options,
      )).toThrow(TypeError);
    }
    expect(reads).toBe(0);
  });

  it('rejects invalid appearance fields through the shared image contract', () => {
    const invalid: readonly Record<string, unknown>[] = [
      { name: null },
      { name: 'bad\u0000name' },
      { altText: false },
      { altText: 'bad\u000Btext' },
      { x: Number.NaN },
      { y: '1' },
      { width: 0 },
      { height: 1.5 },
      { rotation: 21_600_001 },
      { flipHorizontal: 1 },
      { flipVertical: 'false' },
      { sourceRectangle: { left: 60, top: 0, right: 40, bottom: 0 } },
    ];
    for (const options of invalid) {
      expect(() => normalizeEmbeddedSvgImage(
        new Uint8Array([1]),
        new Uint8Array([2]),
        options,
      )).toThrow();
    }
  });
});

describe('embedded SVG picture rendering', () => {
  it('renders escaped canonical paired blip XML and direct appearance state', () => {
    const definition = normalizeEmbeddedSvgImage(
      new Uint8Array([1]),
      new Uint8Array([2]),
      {
        name: 'Vector & <logo>',
        altText: 'Quarterly "result" & <trend>',
        x: -12,
        y: 34,
        width: 56,
        height: 78,
        rotation: 2_700_000,
        flipHorizontal: true,
        flipVertical: true,
        sourceRectangle: { left: 25, top: -10, right: 5, bottom: 0 },
      },
    );
    const source = renderEmbeddedSvgImageXml(
      7,
      definition,
      'rId&fallback',
      'rId<svg',
      'Image 0',
    );
    const { xml, picture } = parsePicture(source);

    expect(source).toContain('name="Vector &amp; &lt;logo&gt;"');
    expect(source).toContain('descr="Quarterly &quot;result&quot; &amp; &lt;trend&gt;"');
    expect(source).toContain('<a:blip r:embed="rId&amp;fallback">');
    expect(source).toContain(`uri="${SVG_EXTENSION_URI}"`);
    expect(source).toContain(`xmlns:asvg="${SVG_NAMESPACE}"`);
    expect(source).toContain('r:embed="rId&lt;svg"');
    expect(source).toContain('rot="2700000" flipH="1" flipV="1"');

    expect(directChildren(picture).map(({ localName }) => localName)).toEqual([
      'nvPicPr',
      'blipFill',
      'spPr',
    ]);
    const blipFill = directChildren(picture)[1]!;
    expect(directChildren(blipFill).map(({ localName }) => localName)).toEqual([
      'blip',
      'srcRect',
      'stretch',
    ]);
    const blip = directChildren(blipFill)[0]!;
    expect(xml.attribute(blip, 'r:embed')?.value).toBe('rId&fallback');
    expect(directChildren(blip).map(({ localName }) => localName)).toEqual(['extLst']);
    const extensionList = directChildren(blip)[0]!;
    const extension = directChildren(extensionList)[0]!;
    expect(extension.localName).toBe('ext');
    expect(xml.attribute(extension, 'uri')?.value).toBe(SVG_EXTENSION_URI);
    const svgBlip = directChildren(extension)[0]!;
    expect(svgBlip.name).toBe('asvg:svgBlip');
    expect(xml.attribute(svgBlip, 'xmlns:asvg')?.value).toBe(SVG_NAMESPACE);
    expect(xml.attribute(svgBlip, 'r:embed')?.value).toBe('rId<svg');
    expect(svgBlip.children).toEqual([]);

    const sourceRectangle = directChildren(blipFill)[1]!;
    expect(['l', 't', 'r', 'b'].map((name) => xml.attribute(sourceRectangle, name)?.value))
      .toEqual(['25000', '-10000', '5000', '0']);
    expect(xml.descendants(picture, 'picLocks')).toHaveLength(1);
    expect(xml.descendants(picture, 'fillRect')).toHaveLength(1);
  });

  it('preserves default and explicit-empty metadata without optional transform XML', () => {
    const omitted = renderEmbeddedSvgImageXml(
      2,
      normalizeEmbeddedSvgImage(new Uint8Array([1]), new Uint8Array([2]), {}),
      'rId1',
      'rId2',
      'Image 0',
    );
    expect(omitted).toContain('name="Image 0" descr="preencoded.svg"');
    expect(omitted).toContain('<a:xfrm>');
    expect(omitted).not.toContain(' rot=');
    expect(omitted).not.toContain(' flipH=');
    expect(omitted).not.toContain(' flipV=');
    expect(omitted).not.toContain('srcRect');

    const empty = renderEmbeddedSvgImageXml(
      3,
      normalizeEmbeddedSvgImage(
        new Uint8Array([1]),
        new Uint8Array([2]),
        { name: '', altText: '' },
      ),
      'rId3',
      'rId4',
      'Image 1',
    );
    expect(empty).toContain('name="" descr=""');
    expect(empty).not.toContain('name="Image 1"');
  });
});

function parsePicture(source: string): { xml: LosslessXmlDocument; picture: XmlElement } {
  const xml = LosslessXmlDocument.parse(
    `<root xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" `
      + `xmlns:r="${RELATIONSHIP_NAMESPACE}">${source}</root>`,
  );
  const picture = xml.elements('pic')[0];
  if (!picture) throw new Error('Rendered fixture has no picture');
  return { xml, picture };
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter((child): child is XmlElement => child.type === 'element');
}
