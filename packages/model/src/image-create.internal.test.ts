import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument, type XmlElement } from '@pptx/lossless-xml';
import type { AddImageOptions, RasterImageContentType } from './image.js';
import {
  normalizeEmbeddedRasterImage,
  renderEmbeddedRasterImageXml,
} from './image-create.internal.js';
import { degrees, inches } from './units.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

describe('embedded raster image normalization', () => {
  it('publishes the closed content-type surface and detached one-inch defaults', () => {
    const contentTypes: readonly RasterImageContentType[] = [
      'image/png',
      'image/jpeg',
      'image/gif',
    ];
    const bytes = new Uint8Array([137, 80, 78, 71]);
    const options: AddImageOptions = { contentType: contentTypes[0]! };
    const normalized = normalizeEmbeddedRasterImage(bytes, options);

    expect(normalized).toEqual({
      bytes: new Uint8Array([137, 80, 78, 71]),
      contentType: 'image/png',
      extension: '.png',
      name: undefined,
      altText: 'preencoded.png',
      x: 0,
      y: 0,
      width: 914_400,
      height: 914_400,
      rotation: 0,
      flipHorizontal: false,
      flipVertical: false,
    });
    expect(normalized.bytes).not.toBe(bytes);
    expect(Object.isFrozen(normalized)).toBe(true);

    bytes.fill(0);
    expect(normalized.bytes).toEqual(new Uint8Array([137, 80, 78, 71]));
  });

  it('maps every content type to a stable extension', () => {
    const cases: readonly [RasterImageContentType, string][] = [
      ['image/png', '.png'],
      ['image/jpeg', '.jpeg'],
      ['image/gif', '.gif'],
    ];
    for (const [contentType, extension] of cases) {
      expect(normalizeEmbeddedRasterImage(
        new Uint8Array([1]),
        { contentType },
      )).toMatchObject({ contentType, extension });
    }
  });

  it('accepts null-prototype options and preserves explicit direct values', () => {
    const options = Object.create(null) as Record<string, unknown>;
    options.contentType = 'image/jpeg';
    options.name = '';
    options.altText = '';
    options.x = -12;
    options.y = 34;
    options.width = 56;
    options.height = 78;
    options.rotation = -2_700_000;
    options.flipHorizontal = true;
    options.flipVertical = true;

    const normalized = normalizeEmbeddedRasterImage(new Uint8Array([1, 2]), options);
    options.name = 'changed';
    options.altText = 'changed';
    options.x = 999;

    expect(normalized).toEqual({
      bytes: new Uint8Array([1, 2]),
      contentType: 'image/jpeg',
      extension: '.jpeg',
      name: '',
      altText: '',
      x: -12,
      y: 34,
      width: 56,
      height: 78,
      rotation: -2_700_000,
      flipHorizontal: true,
      flipVertical: true,
    });
  });

  it('accepts ergonomic converter output through the public options type', () => {
    const options: AddImageOptions = {
      contentType: 'image/gif',
      x: inches(1),
      y: inches(2),
      width: inches(3),
      height: inches(4),
      rotation: degrees(45),
      flipHorizontal: true,
    };

    expect(normalizeEmbeddedRasterImage(new Uint8Array([71]), options)).toMatchObject({
      x: 914_400,
      y: 1_828_800,
      width: 2_743_200,
      height: 3_657_600,
      rotation: 2_700_000,
      flipHorizontal: true,
      flipVertical: false,
    });
  });

  it('rejects non-byte and empty payloads', () => {
    const options = { contentType: 'image/png' };
    for (const value of [undefined, null, [], [1], new ArrayBuffer(1), 'png', 1]) {
      expect(() => normalizeEmbeddedRasterImage(value, options), String(value)).toThrow(TypeError);
    }
    expect(() => normalizeEmbeddedRasterImage(new Uint8Array(), options)).toThrow(RangeError);
  });

  it('rejects unsafe option containers and descriptors without reading accessors', () => {
    class Options {
      contentType = 'image/png';
    }
    const inherited = Object.create({ contentType: 'image/png' });
    const symbol = { contentType: 'image/png', [Symbol('unsafe')]: true };
    const unknown = { contentType: 'image/png', unsupported: undefined };
    let reads = 0;
    const accessor = Object.defineProperty({ contentType: 'image/png' }, 'name', {
      get: () => {
        reads += 1;
        return 'unsafe';
      },
    });

    for (const options of [null, [], new Options(), inherited, symbol, unknown, accessor]) {
      expect(() => normalizeEmbeddedRasterImage(new Uint8Array([1]), options)).toThrow(TypeError);
    }
    expect(reads).toBe(0);
  });

  it('rejects missing and unsupported content types', () => {
    for (const options of [
      {},
      { contentType: undefined },
      { contentType: null },
      { contentType: 'image/svg+xml' },
      { contentType: 'image/jpg' },
      { contentType: 1 },
    ]) {
      expect(() => normalizeEmbeddedRasterImage(new Uint8Array([1]), options)).toThrow(TypeError);
    }
  });

  it('rejects invalid names and alt text', () => {
    for (const [property, value] of [
      ['name', null],
      ['name', 1],
      ['name', 'bad\u0000name'],
      ['altText', false],
      ['altText', 'bad\u000Btext'],
    ] as const) {
      expect(() => normalizeEmbeddedRasterImage(
        new Uint8Array([1]),
        { contentType: 'image/png', [property]: value },
      )).toThrow(TypeError);
    }
  });

  it('rejects unsafe transforms', () => {
    const invalid: readonly [string, unknown][] = [
      ['x', Number.NaN],
      ['x', Number.POSITIVE_INFINITY],
      ['x', 1.5],
      ['x', Number.MAX_SAFE_INTEGER + 1],
      ['y', '1'],
      ['width', 0],
      ['width', -1],
      ['height', 0],
      ['height', 1.5],
      ['rotation', -21_600_001],
      ['rotation', 21_600_001],
      ['rotation', 0.5],
      ['flipHorizontal', 1],
      ['flipVertical', 'false'],
    ];
    for (const [property, value] of invalid) {
      expect(() => normalizeEmbeddedRasterImage(
        new Uint8Array([1]),
        { contentType: 'image/png', [property]: value },
      ), `${property}=${String(value)}`).toThrow();
    }
  });
});

describe('embedded raster picture rendering', () => {
  it('renders escaped canonical picture XML with direct transform state', () => {
    const definition = normalizeEmbeddedRasterImage(new Uint8Array([1]), {
      contentType: 'image/png',
      name: 'Revenue & <logo>',
      altText: 'Quarterly "result" & <trend>',
      x: -12,
      y: 34,
      width: 56,
      height: 78,
      rotation: 2_700_000,
      flipHorizontal: true,
      flipVertical: true,
    });
    const source = renderEmbeddedRasterImageXml(7, definition, 'rId4', 'Image 0');
    const { xml, picture } = parsePicture(source);

    expect(source).toContain('id="7"');
    expect(source).toContain('name="Revenue &amp; &lt;logo&gt;"');
    expect(source).toContain('descr="Quarterly &quot;result&quot; &amp; &lt;trend&gt;"');
    expect(source).toContain('r:embed="rId4"');
    expect(source).toContain('rot="2700000" flipH="1" flipV="1"');
    expect(source).toContain('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>');

    expect(directChildren(picture).map(({ localName }) => localName)).toEqual([
      'nvPicPr',
      'blipFill',
      'spPr',
    ]);
    const nonVisual = directChildren(picture)[0]!;
    expect(directChildren(nonVisual).map(({ localName }) => localName)).toEqual([
      'cNvPr',
      'cNvPicPr',
      'nvPr',
    ]);
    const locks = xml.descendants(picture, 'picLocks');
    expect(locks).toHaveLength(1);
    expect(xml.attribute(locks[0]!, 'noChangeAspect')?.value).toBe('1');
    const blip = xml.descendants(picture, 'blip');
    expect(blip).toHaveLength(1);
    expect(xml.attribute(blip[0]!, 'r:embed')?.value).toBe('rId4');
    expect(xml.descendants(picture, 'stretch')).toHaveLength(1);
    expect(xml.descendants(picture, 'fillRect')).toHaveLength(1);
    expect(xml.descendants(picture, 'off')).toHaveLength(1);
    expect(xml.descendants(picture, 'ext')).toHaveLength(1);
  });

  it('uses the default name while preserving explicit empty strings and default flags', () => {
    const omitted = normalizeEmbeddedRasterImage(
      new Uint8Array([1]),
      { contentType: 'image/png' },
    );
    const omittedSource = renderEmbeddedRasterImageXml(2, omitted, 'rId1', 'Image 0');
    expect(omittedSource).toContain('name="Image 0" descr="preencoded.png"');
    expect(omittedSource).toContain('<a:xfrm>');
    expect(omittedSource).not.toContain(' rot=');
    expect(omittedSource).not.toContain(' flipH=');
    expect(omittedSource).not.toContain(' flipV=');

    const empty = normalizeEmbeddedRasterImage(
      new Uint8Array([1]),
      { contentType: 'image/gif', name: '', altText: '' },
    );
    const emptySource = renderEmbeddedRasterImageXml(3, empty, 'rId2', 'Image 1');
    expect(emptySource).toContain('name="" descr=""');
    expect(emptySource).not.toContain('name="Image 1"');
  });
});

function parsePicture(source: string): { xml: LosslessXmlDocument; picture: XmlElement } {
  const xml = LosslessXmlDocument.parse(
    `<root xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" ` +
      `xmlns:r="${RELATIONSHIP_NAMESPACE}">${source}</root>`,
  );
  const picture = xml.elements('pic')[0];
  if (!picture) throw new Error('Rendered fixture has no picture');
  return { xml, picture };
}

function directChildren(element: XmlElement): XmlElement[] {
  return element.children.filter((child): child is XmlElement => child.type === 'element');
}
