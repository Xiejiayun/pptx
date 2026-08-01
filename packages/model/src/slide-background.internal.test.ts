import { describe, expect, it } from 'vitest';
import { OpcPackage, relativeRelationshipTarget } from '@pptx/opc';
import type { SlideBackground } from './slide-background.js';
import {
  normalizeSlideBackground,
  readSlideBackground,
  replaceSlideBackground,
  slideBackgroundMediaTargets,
} from './slide-background.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const IMAGE_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}/image`;
const SLIDE_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const SLIDE_URI = '/ppt/slides/slide1.xml';

const BACKGROUND_OWNER_FIXTURES = [
  {
    kind: 'layout' as const,
    root: 'sldLayout',
    uri: '/ppt/slideLayouts/slideLayout1.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
  },
  {
    kind: 'master' as const,
    root: 'sldMaster',
    uri: '/ppt/slideMasters/slideMaster1.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
  },
] as const;

describe('slide background normalization', () => {
  it('normalizes simple fills and detached raster image bytes', () => {
    expect(normalizeSlideBackground(undefined)).toBeUndefined();
    const none: SlideBackground = { kind: 'none' };
    expect(normalizeSlideBackground(none)).toEqual({ kind: 'none' });
    expect(normalizeSlideBackground({
      kind: 'solid',
      color: { kind: 'srgb', value: '#ff3399' },
      transparency: 50.0004,
    })).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF3399' },
      transparency: 50,
    });

    const bytes = Uint8Array.of(1, 2, 3);
    const normalized = normalizeSlideBackground({
      kind: 'image',
      contentType: 'image/png',
      bytes,
    });
    bytes[0] = 9;
    expect(normalized).toEqual({
      kind: 'image',
      contentType: 'image/png',
      bytes: Uint8Array.of(1, 2, 3),
    });
    expect(Object.isFrozen(normalized)).toBe(true);
    if (normalized?.kind === 'image') expect(normalized.bytes).not.toBe(bytes);
  });

  it('normalizes gradients into deeply detached frozen values', () => {
    const source = { kind: 'system', value: 'windowText', lastColor: '112233' };
    const transforms = [{ kind: 'tint', value: 20_000 }];
    const stops = [
      { offset: 0, color: '#ff0000' },
      {
        offset: 1,
        alpha: 0.5,
        color: { source, alpha: 0.75, transforms },
      },
    ];
    const rectangle = { left: 0, top: 0.1, right: 0.2, bottom: 0.3 };
    const input = {
      kind: 'path-gradient',
      path: 'circle',
      rotateWithShape: false,
      fillRectangle: rectangle,
      stops,
    };

    const normalized = normalizeSlideBackground(input);
    source.value = 'window';
    transforms[0]!.value = 1;
    stops[0]!.offset = 0.5;
    rectangle.left = 0.5;

    expect(normalized).toEqual({
      kind: 'path-gradient',
      path: 'circle',
      rotateWithShape: false,
      fillRectangle: { left: 0, top: 0.1, right: 0.2, bottom: 0.3 },
      stops: [
        { offset: 0, color: 'FF0000' },
        {
          offset: 1,
          alpha: 0.5,
          color: {
            source: { kind: 'system', value: 'windowText', lastColor: '112233' },
            alpha: 0.75,
            transforms: [{ kind: 'tint', value: 20_000 }],
          },
        },
      ],
    });
    expect(Object.isFrozen(normalized)).toBe(true);
    if (normalized?.kind !== 'path-gradient') throw new Error('Expected path gradient');
    expect(Object.isFrozen(normalized.stops)).toBe(true);
    expect(Object.isFrozen(normalized.stops[1])).toBe(true);
    const color = normalized.stops[1]?.color;
    if (typeof color === 'string' || color === undefined) throw new Error('Expected OOXML color');
    expect(Object.isFrozen(color)).toBe(true);
    expect(Object.isFrozen(color.source)).toBe(true);
    expect(Object.isFrozen(color.transforms)).toBe(true);
    expect(Object.isFrozen(color.transforms[0])).toBe(true);
    expect(Object.isFrozen(normalized.fillRectangle)).toBe(true);
  });

  it('accepts null-prototype own-data objects', () => {
    const color = Object.assign(Object.create(null), {
      kind: 'scheme',
      value: 'accent2',
    });
    const fill = Object.assign(Object.create(null), {
      kind: 'solid',
      color,
      transparency: 0,
    });
    expect(normalizeSlideBackground(fill)).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent2' },
      transparency: 0,
    });
  });

  it('rejects unsafe containers and accessors without invoking them', () => {
    class Background {
      kind = 'none';
    }
    for (const value of [null, false, 1, 'none', [], new Date(), new Background()]) {
      expect(() => normalizeSlideBackground(value)).toThrow(TypeError);
    }
    expect(() => normalizeSlideBackground(Object.create({ kind: 'none' })))
      .toThrow(TypeError);
    expect(() => normalizeSlideBackground({ kind: 'none', extra: true }))
      .toThrow(/unsupported property extra/);
    expect(() => normalizeSlideBackground({ kind: 'none', [Symbol('unsafe')]: true }))
      .toThrow(/unsupported property/);

    let calls = 0;
    const kindAccessor = Object.defineProperty({}, 'kind', {
      enumerable: true,
      get() {
        calls += 1;
        return 'none';
      },
    });
    const bytesAccessor = Object.defineProperty({
      kind: 'image',
      contentType: 'image/png',
    }, 'bytes', {
      enumerable: true,
      get() {
        calls += 1;
        return Uint8Array.of(1);
      },
    });
    for (const value of [kindAccessor, bytesAccessor]) {
      expect(() => normalizeSlideBackground(value)).toThrow(/data property/);
    }
    expect(calls).toBe(0);
  });

  it('rejects invalid raster and gradient states', () => {
    const sparseStops = new Array(2);
    sparseStops[1] = { offset: 1, color: '000000' };
    const sparseTransforms = new Array(2);
    sparseTransforms[1] = { kind: 'tint', value: 20_000 };
    for (const value of [
      { kind: 'image', contentType: 'image/png', bytes: [] },
      { kind: 'image', contentType: 'image/svg+xml', bytes: Uint8Array.of(1) },
      { kind: 'image', contentType: 'image/png', bytes: new Uint8Array() },
      { kind: 'linear-gradient', angle: 0, stops: [] },
      { kind: 'linear-gradient', angle: Number.NaN, stops: validStops() },
      { kind: 'linear-gradient', angle: 0, scaled: 1, stops: validStops() },
      { kind: 'linear-gradient', angle: 0, flip: 'z', stops: validStops() },
      { kind: 'path-gradient', path: 'ellipse', stops: validStops() },
      { kind: 'linear-gradient', angle: 0, stops: sparseStops },
      {
        kind: 'path-gradient',
        path: 'circle',
        fillRectangle: { left: 0, top: 0, right: Number.POSITIVE_INFINITY, bottom: 0 },
        stops: validStops(),
      },
      {
        kind: 'linear-gradient',
        angle: 0,
        stops: [{ offset: 0, color: 'FFFFFF' }, { offset: 1.1, color: '000000' }],
      },
      {
        kind: 'linear-gradient',
        angle: 0,
        stops: [{ offset: 0, color: 'FFFFFF' }, { offset: 1, color: '000000', alpha: -0.1 }],
      },
      {
        kind: 'linear-gradient',
        angle: 0,
        stops: [{ offset: 0, color: 'FFF' }, { offset: 1, color: '000000' }],
      },
      {
        kind: 'linear-gradient',
        angle: 0,
        stops: [
          { offset: 0, color: 'FFFFFF' },
          {
            offset: 1,
            color: {
              source: { kind: 'scheme', value: '' },
              alpha: 1,
              transforms: [],
            },
          },
        ],
      },
      {
        kind: 'linear-gradient',
        angle: 0,
        stops: [
          { offset: 0, color: 'FFFFFF' },
          {
            offset: 1,
            color: {
              source: { kind: 'scheme', value: 'accent1' },
              alpha: 1,
              transforms: sparseTransforms,
            },
          },
        ],
      },
    ]) {
      expect(() => normalizeSlideBackground(value), JSON.stringify(value)).toThrow();
    }
  });
});

describe('strict direct slide background reader', () => {
  it('reads and edits strict direct layout master backgrounds', () => {
    for (const owner of BACKGROUND_OWNER_FIXTURES) {
      const pkg = OpcPackage.create();
      pkg.setPart(
        owner.uri,
        `<p:${owner.root} xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}">`
          + '<p:cSld><p:spTree/></p:cSld>'
          + `</p:${owner.root}>`,
        owner.contentType,
      );
      expect(readSlideBackground(pkg, owner.uri, owner.kind)).toBeUndefined();

      replaceSlideBackground(pkg, owner.uri, {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 20,
      }, owner.kind);
      expect(readSlideBackground(pkg, owner.uri, owner.kind)).toEqual({
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 20,
      });
      expect(readSlideBackground(pkg, owner.uri)).toBeUndefined();

      replaceSlideBackground(pkg, owner.uri, { kind: 'none' }, owner.kind);
      expect(readSlideBackground(pkg, owner.uri, owner.kind)).toEqual({ kind: 'none' });
      replaceSlideBackground(pkg, owner.uri, undefined, owner.kind);
      expect(readSlideBackground(pkg, owner.uri, owner.kind)).toBeUndefined();
    }
  });

  it.each([
    ['no fill', '<a:noFill/>', { kind: 'none' }],
    [
      'sRGB solid',
      '<a:solidFill><a:srgbClr val="ff3399"><a:alpha val="50000"/></a:srgbClr></a:solidFill>',
      {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF3399' },
        transparency: 50,
      },
    ],
    [
      'scheme solid',
      '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>',
      { kind: 'solid', color: { kind: 'scheme', value: 'accent1' } },
    ],
    [
      'alternate-prefix solid',
      `<d:solidFill xmlns:d="${DRAWING_NAMESPACE}"><a:srgbClr val="112233">`
        + `<d:alpha val="75000"/></a:srgbClr></d:solidFill>`,
      {
        kind: 'solid',
        color: { kind: 'srgb', value: '112233' },
        transparency: 25,
      },
    ],
  ] as const)('reads %s without mutating the package', (_name, fill, expected) => {
    const pkg = backgroundPackage(`<p:bg><p:bgPr>${fill}<a:effectLst/></p:bgPr></p:bg>`);
    const before = snapshot(pkg);
    expect(readSlideBackground(pkg, SLIDE_URI)).toEqual(expected);
    expect(snapshot(pkg)).toEqual(before);
  });

  it('reads direct linear and path gradients into detached frozen values', () => {
    const linear = backgroundPackage(
      '<p:bg><p:bgPr><a:gradFill flip="x" rotWithShape="false"><a:gsLst>'
        + '<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>'
        + '<a:gs pos="100000"><a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr></a:gs>'
        + '</a:gsLst><a:lin ang="2700000" scaled="false"/></a:gradFill>'
        + '<a:effectLst/></p:bgPr></p:bg>',
    );
    const result = readSlideBackground(linear, SLIDE_URI);
    expect(result).toMatchObject({
      kind: 'linear-gradient',
      angle: 45,
      scaled: false,
      rotateWithShape: false,
      flip: 'x',
      stops: [
        { offset: 0, alpha: 1 },
        { offset: 1, alpha: 0.5 },
      ],
    });
    expect(Object.isFrozen(result)).toBe(true);
    if (result?.kind !== 'linear-gradient') throw new Error('Expected linear gradient');
    expect(Object.isFrozen(result.stops)).toBe(true);
    expect(Object.isFrozen(result.stops[0])).toBe(true);

    const path = backgroundPackage(
      '<p:bg><p:bgPr><a:gradFill><a:gsLst>'
        + '<a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>'
        + '<a:gs pos="100000"><a:srgbClr val="000000"/></a:gs>'
        + '</a:gsLst><a:path path="circle"><a:fillToRect l="10000" t="20000" r="30000" b="40000"/></a:path>'
        + '</a:gradFill></p:bgPr></p:bg>',
    );
    expect(readSlideBackground(path, SLIDE_URI)).toMatchObject({
      kind: 'path-gradient',
      path: 'circle',
      fillRectangle: { left: 0.1, top: 0.2, right: 0.3, bottom: 0.4 },
    });
  });

  it.each([
    ['image/png', '.png'],
    ['image/jpeg', '.jpeg'],
    ['image/gif', '.gif'],
  ] as const)('reads detached %s image payloads', (contentType, extension) => {
    const target = `/ppt/media/background1${extension}`;
    const pkg = backgroundPackage(
      '<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rId7"/>'
        + '<a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg>',
      {
        parts: [{ uri: target, contentType, bytes: Uint8Array.of(1, 2, 3) }],
        relationships: [{
          id: 'rId7',
          type: IMAGE_RELATIONSHIP,
          target,
        }],
      },
    );
    const before = snapshot(pkg);
    const first = readSlideBackground(pkg, SLIDE_URI);
    const second = readSlideBackground(pkg, SLIDE_URI);
    expect(first).toEqual({
      kind: 'image',
      contentType,
      bytes: Uint8Array.of(1, 2, 3),
    });
    expect(Object.isFrozen(first)).toBe(true);
    if (first?.kind !== 'image' || second?.kind !== 'image') {
      throw new Error('Expected image backgrounds');
    }
    expect(first.bytes).not.toBe(second.bytes);
    first.bytes[0] = 9;
    expect(second.bytes).toEqual(Uint8Array.of(1, 2, 3));
    expect(pkg.requirePart(target).bytes).toEqual(Uint8Array.of(1, 2, 3));
    expect(snapshot(pkg)).toEqual(before);
  });

  it('returns undefined when the direct background is absent', () => {
    const pkg = backgroundPackage('');
    expect(readSlideBackground(pkg, SLIDE_URI)).toBeUndefined();
  });

  it.each([
    ['background reference', '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'],
    ['empty properties', '<p:bg><p:bgPr><a:effectLst/></p:bgPr></p:bg>'],
    ['pattern fill', '<p:bg><p:bgPr><a:pattFill prst="pct5"/></p:bgPr></p:bg>'],
    ['group fill', '<p:bg><p:bgPr><a:grpFill/></p:bgPr></p:bg>'],
    ['multiple choices', '<p:bg><p:bgPr><a:noFill/><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg>'],
    ['nested fill trap', '<p:bg><p:bgPr><a:effectLst><a:noFill/></a:effectLst></p:bgPr></p:bg>'],
    ['wrong fill namespace', '<p:bg><p:bgPr><x:noFill xmlns:x="urn:not-drawing"/></p:bgPr></p:bg>'],
    ['malformed solid', '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFF"/></a:solidFill></p:bgPr></p:bg>'],
    [
      'malformed gradient color',
      '<p:bg><p:bgPr><a:gradFill><a:gsLst>'
        + '<a:gs pos="0"><a:srgbClr/></a:gs>'
        + '<a:gs pos="100000"><a:srgbClr val="000000"/></a:gs>'
        + '</a:gsLst><a:lin ang="0"/></a:gradFill></p:bgPr></p:bg>',
    ],
    [
      'ambiguous gradient mode',
      '<p:bg><p:bgPr><a:gradFill><a:gsLst>'
        + '<a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>'
        + '<a:gs pos="100000"><a:srgbClr val="000000"/></a:gs>'
        + '</a:gsLst><a:lin ang="0"/><a:path path="circle"/></a:gradFill></p:bgPr></p:bg>',
    ],
    [
      'invalid gradient attribute',
      '<p:bg><p:bgPr><a:gradFill flip="z"><a:gsLst>'
        + '<a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs>'
        + '<a:gs pos="100000"><a:srgbClr val="000000"/></a:gs>'
        + '</a:gsLst><a:lin ang="0" scaled="maybe"/></a:gradFill></p:bgPr></p:bg>',
    ],
    ['duplicate properties', '<p:bg><p:bgPr><a:noFill/></p:bgPr><p:bgPr><a:noFill/></p:bgPr></p:bg>'],
  ] as const)('does not project unsafe %s state', (_name, background) => {
    const pkg = backgroundPackage(background);
    const before = snapshot(pkg);
    expect(readSlideBackground(pkg, SLIDE_URI)).toBeUndefined();
    expect(snapshot(pkg)).toEqual(before);
  });

  it('rejects descendant and duplicate direct-chain traps', () => {
    const descendant = packageFromSlide(
      `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" `
        + `xmlns:r="${RELATIONSHIP_NAMESPACE}"><p:extLst><p:ext><p:cSld>`
        + '<p:bg><p:bgPr><a:noFill/></p:bgPr></p:bg></p:cSld></p:ext></p:extLst>'
        + '<p:cSld><p:spTree/></p:cSld></p:sld>',
    );
    expect(readSlideBackground(descendant, SLIDE_URI)).toBeUndefined();

    const duplicate = packageFromSlide(
      `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" `
        + `xmlns:r="${RELATIONSHIP_NAMESPACE}"><p:cSld><p:bg><p:bgPr><a:noFill/>`
        + '</p:bgPr></p:bg><p:spTree/></p:cSld><p:cSld><p:spTree/></p:cSld></p:sld>',
    );
    expect(readSlideBackground(duplicate, SLIDE_URI)).toBeUndefined();
  });

  it.each([
    ['unqualified embed', '<a:blip embed="rId7"/>'],
    ['wrong namespace embed', '<a:blip x:embed="rId7" xmlns:x="urn:not-relationship"/>'],
    ['qualified and unqualified embed', '<a:blip r:embed="rId7" embed="rId7"/>'],
    ['duplicate blips', '<a:blip r:embed="rId7"/><a:blip r:embed="rId7"/>'],
  ] as const)('rejects image state with %s', (_name, blip) => {
    const pkg = backgroundPackage(
      `<p:bg><p:bgPr><a:blipFill>${blip}<a:stretch><a:fillRect/></a:stretch>`
        + '</a:blipFill></p:bgPr></p:bg>',
      {
        parts: [{
          uri: '/ppt/media/background1.png',
          contentType: 'image/png',
          bytes: Uint8Array.of(1),
        }],
        relationships: [{
          id: 'rId7',
          type: IMAGE_RELATIONSHIP,
          target: '/ppt/media/background1.png',
        }],
      },
    );
    expect(readSlideBackground(pkg, SLIDE_URI)).toBeUndefined();
  });

  it.each([
    ['external', IMAGE_RELATIONSHIP, 'https://example.com/image.png', 'External'],
    ['wrong relationship type', `${RELATIONSHIP_NAMESPACE}/audio`, '/ppt/media/background1.png', 'Internal'],
    ['dangling target', IMAGE_RELATIONSHIP, '/ppt/media/missing.png', 'Internal'],
  ] as const)('rejects %s image relationships', (_name, type, target, targetMode) => {
    const pkg = backgroundPackage(
      '<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rId7"/></a:blipFill></p:bgPr></p:bg>',
      {
        parts: target === '/ppt/media/background1.png'
          ? [{
              uri: target,
              contentType: 'image/png',
              bytes: Uint8Array.of(1),
            }]
          : [],
        relationships: [{ id: 'rId7', type, target, targetMode }],
      },
    );
    expect(readSlideBackground(pkg, SLIDE_URI)).toBeUndefined();
  });

  it('rejects duplicate relationship ids and unsupported target MIME', () => {
    const target = '/ppt/media/background1.svg';
    const pkg = backgroundPackage(
      '<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rId7"/></a:blipFill></p:bgPr></p:bg>',
      {
        parts: [{
          uri: target,
          contentType: 'image/svg+xml',
          bytes: Uint8Array.of(1),
        }],
        relationships: [
          { id: 'rId7', type: IMAGE_RELATIONSHIP, target },
          { id: 'rId7', type: IMAGE_RELATIONSHIP, target },
        ],
      },
    );
    expect(readSlideBackground(pkg, SLIDE_URI)).toBeUndefined();

    const single = backgroundPackage(
      '<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rId7"/></a:blipFill></p:bgPr></p:bg>',
      {
        parts: [{
          uri: target,
          contentType: 'image/svg+xml',
          bytes: Uint8Array.of(1),
        }],
        relationships: [{ id: 'rId7', type: IMAGE_RELATIONSHIP, target }],
      },
    );
    expect(readSlideBackground(single, SLIDE_URI)).toBeUndefined();
  });
});

describe('layout master background owner parity', () => {
  it.each(BACKGROUND_OWNER_FIXTURES)(
    'reads, edits, and cleans $kind background values',
    (owner) => {
      const pkg = ownerBackgroundPackage(owner, '');

      replaceSlideBackground(pkg, owner.uri, { kind: 'none' }, owner.kind);
      expect(readSlideBackground(pkg, owner.uri, owner.kind)).toEqual({ kind: 'none' });

      const solid = {
        kind: 'solid',
        color: { kind: 'srgb', value: 'FF3399' },
        transparency: 50,
      } as const;
      replaceSlideBackground(pkg, owner.uri, solid, owner.kind);
      expect(readSlideBackground(pkg, owner.uri, owner.kind)).toEqual(solid);
      const solidSnapshot = snapshot(pkg, owner.uri);
      replaceSlideBackground(pkg, owner.uri, solid, owner.kind);
      expect(snapshot(pkg, owner.uri)).toEqual(solidSnapshot);

      replaceSlideBackground(pkg, owner.uri, {
        kind: 'linear-gradient',
        angle: 45,
        scaled: false,
        stops: validStops(),
      }, owner.kind);
      expect(readSlideBackground(pkg, owner.uri, owner.kind)).toMatchObject({
        kind: 'linear-gradient',
        angle: 45,
        scaled: false,
      });

      replaceSlideBackground(pkg, owner.uri, {
        kind: 'path-gradient',
        path: 'circle',
        fillRectangle: { left: 0.1, top: 0.2, right: 0.3, bottom: 0.4 },
        stops: validStops(),
      }, owner.kind);
      expect(readSlideBackground(pkg, owner.uri, owner.kind)).toMatchObject({
        kind: 'path-gradient',
        path: 'circle',
        fillRectangle: { left: 0.1, top: 0.2, right: 0.3, bottom: 0.4 },
      });

      const bytes = Uint8Array.of(1, 2, 3);
      replaceSlideBackground(pkg, owner.uri, {
        kind: 'image',
        contentType: 'image/png',
        bytes,
      }, owner.kind);
      bytes[0] = 9;
      const target = slideBackgroundMediaTargets(pkg, owner.uri, owner.kind)[0]!;
      expect(target).toMatch(/\/ppt\/media\/background\d+\.png$/);
      expect(readSlideBackground(pkg, owner.uri, owner.kind)).toEqual({
        kind: 'image',
        contentType: 'image/png',
        bytes: Uint8Array.of(1, 2, 3),
      });
      expect(pkg.requirePart(target).bytes).toEqual(Uint8Array.of(1, 2, 3));

      const imageSnapshot = snapshot(pkg, owner.uri);
      replaceSlideBackground(pkg, owner.uri, {
        kind: 'image',
        contentType: 'image/png',
        bytes: Uint8Array.of(1, 2, 3),
      }, owner.kind);
      expect(snapshot(pkg, owner.uri)).toEqual(imageSnapshot);

      replaceSlideBackground(pkg, owner.uri, undefined, owner.kind);
      expect(readSlideBackground(pkg, owner.uri, owner.kind)).toBeUndefined();
      expect(pkg.hasPart(target)).toBe(false);
      expect(slideBackgroundMediaTargets(pkg, owner.uri, owner.kind)).toEqual([]);

      const clearSnapshot = snapshot(pkg, owner.uri);
      replaceSlideBackground(pkg, owner.uri, undefined, owner.kind);
      expect(snapshot(pkg, owner.uri)).toEqual(clearSnapshot);
      expect(() => pkg.transaction(() => {
        replaceSlideBackground(pkg, owner.uri, {
          kind: 'solid',
          color: { kind: 'scheme', value: 'accent1' },
        }, owner.kind);
        throw new Error('rollback owner background');
      })).toThrow('rollback owner background');
      expect(snapshot(pkg, owner.uri)).toEqual(clearSnapshot);
    },
  );

  it.each(BACKGROUND_OWNER_FIXTURES)(
    'preserves or replaces ambiguous $kind background state atomically',
    (owner) => {
      const reference = ownerBackgroundPackage(
        owner,
        '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>',
      );
      const referenceBefore = snapshot(reference, owner.uri);
      expect(readSlideBackground(reference, owner.uri, owner.kind)).toBeUndefined();
      expect(snapshot(reference, owner.uri)).toEqual(referenceBefore);
      replaceSlideBackground(reference, owner.uri, { kind: 'none' }, owner.kind);
      expect(readSlideBackground(reference, owner.uri, owner.kind)).toEqual({ kind: 'none' });
      expect((ownerXml(reference, owner).match(/<p:bg>/g) ?? [])).toHaveLength(1);

      const opaque = ownerBackgroundPackage(
        owner,
        '<p:bg><p:bgPr><a:pattFill prst="pct5"/></p:bgPr></p:bg>',
      );
      replaceSlideBackground(opaque, owner.uri, {
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 20,
      }, owner.kind);
      expect(readSlideBackground(opaque, owner.uri, owner.kind)).toEqual({
        kind: 'solid',
        color: { kind: 'scheme', value: 'accent1' },
        transparency: 20,
      });

      const wrongNamespace = ownerBackgroundPackage(
        owner,
        '<x:bg><x:keep/></x:bg>',
      );
      expect(readSlideBackground(wrongNamespace, owner.uri, owner.kind)).toBeUndefined();
      replaceSlideBackground(wrongNamespace, owner.uri, { kind: 'none' }, owner.kind);
      expect(ownerXml(wrongNamespace, owner)).toContain('<x:bg><x:keep/></x:bg>');
      expect(readSlideBackground(wrongNamespace, owner.uri, owner.kind)).toEqual({ kind: 'none' });

      const duplicateCommonSlide = packageFromBackgroundOwner(
        owner,
        `<p:${owner.root} xmlns:p="${PRESENTATION_NAMESPACE}" `
          + `xmlns:a="${DRAWING_NAMESPACE}"><p:cSld><p:spTree/></p:cSld>`
          + `<p:cSld><p:spTree/></p:cSld></p:${owner.root}>`,
      );
      const duplicateBefore = snapshot(duplicateCommonSlide, owner.uri);
      expect(readSlideBackground(duplicateCommonSlide, owner.uri, owner.kind)).toBeUndefined();
      expect(() => replaceSlideBackground(
        duplicateCommonSlide,
        owner.uri,
        { kind: 'none' },
        owner.kind,
      )).toThrow(/editable cSld/);
      expect(snapshot(duplicateCommonSlide, owner.uri)).toEqual(duplicateBefore);
    },
  );

  it.each(BACKGROUND_OWNER_FIXTURES)(
    'copy-on-writes shared $kind background media and garbage-collects only its copy',
    (owner) => {
      const target = '/ppt/media/shared-owner.png';
      const pkg = packageFromBackgroundOwner(
        owner,
        `<p:${owner.root} xmlns:p="${PRESENTATION_NAMESPACE}" `
          + `xmlns:a="${DRAWING_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}">`
          + '<p:cSld><p:bg><p:bgPr><a:blipFill><a:blip r:embed="rId7"/>'
          + '</a:blipFill></p:bgPr></p:bg><p:spTree><p:pic><p:blipFill>'
          + '<a:blip r:embed="rId8"/></p:blipFill></p:pic></p:spTree></p:cSld>'
          + `</p:${owner.root}>`,
        {
          parts: [{ uri: target, contentType: 'image/png', bytes: Uint8Array.of(1) }],
          relationships: [
            { id: 'rId7', type: IMAGE_RELATIONSHIP, target },
            { id: 'rId8', type: IMAGE_RELATIONSHIP, target },
          ],
        },
      );

      replaceSlideBackground(pkg, owner.uri, {
        kind: 'image',
        contentType: 'image/png',
        bytes: Uint8Array.of(2),
      }, owner.kind);
      const backgroundRelationship = pkg.relationships(owner.uri).find(({ id }) =>
        id === 'rId7')!;
      expect(backgroundRelationship.resolvedTarget).not.toBe(target);
      expect(pkg.relationships(owner.uri).find(({ id }) => id === 'rId8')?.resolvedTarget)
        .toBe(target);
      expect(pkg.requirePart(target).bytes).toEqual(Uint8Array.of(1));
      expect(pkg.requirePart(backgroundRelationship.resolvedTarget!).bytes)
        .toEqual(Uint8Array.of(2));

      replaceSlideBackground(pkg, owner.uri, undefined, owner.kind);
      expect(pkg.hasPart(backgroundRelationship.resolvedTarget!)).toBe(false);
      expect(pkg.hasPart(target)).toBe(true);
      expect(pkg.relationships(owner.uri).map(({ id }) => id)).toEqual(['rId8']);
    },
  );
});

describe('atomic non-image slide background editing', () => {
  it('creates none, solid, and gradient backgrounds before the shape tree and clears inheritance', () => {
    const pkg = backgroundPackage('');

    replaceSlideBackground(pkg, SLIDE_URI, { kind: 'none' });
    expect(readSlideBackground(pkg, SLIDE_URI)).toEqual({ kind: 'none' });
    expect(slideXml(pkg)).toContain(
      '<p:cSld><p:bg><p:bgPr><a:noFill/><a:effectLst/></p:bgPr></p:bg><p:spTree/>',
    );

    replaceSlideBackground(pkg, SLIDE_URI, {
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 25,
    });
    expect(readSlideBackground(pkg, SLIDE_URI)).toEqual({
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
      transparency: 25,
    });
    expect(slideXml(pkg)).toContain(
      '<a:solidFill><a:schemeClr val="accent1"><a:alpha val="75000"/>'
        + '</a:schemeClr></a:solidFill>',
    );

    replaceSlideBackground(pkg, SLIDE_URI, {
      kind: 'linear-gradient',
      angle: 45,
      scaled: false,
      rotateWithShape: false,
      flip: 'x',
      stops: validStops(),
    });
    expect(readSlideBackground(pkg, SLIDE_URI)).toMatchObject({
      kind: 'linear-gradient',
      angle: 45,
      scaled: false,
      rotateWithShape: false,
      flip: 'x',
    });
    expect(slideXml(pkg)).toContain(
      '<a:gradFill rotWithShape="0" flip="x"><a:gsLst>',
    );

    replaceSlideBackground(pkg, SLIDE_URI, undefined);
    expect(readSlideBackground(pkg, SLIDE_URI)).toBeUndefined();
    expect(slideXml(pkg)).not.toContain('<p:bg>');
    expect(slideXml(pkg)).toContain('<p:cSld><p:spTree/></p:cSld>');
  });

  it('patches a safe fill locally while preserving owned siblings and unrelated XML', () => {
    const pkg = packageFromSlide(
      `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" `
        + `xmlns:r="${RELATIONSHIP_NAMESPACE}" xmlns:x="urn:opaque"><p:cSld custom="keep">`
        + '<p:bg><p:bgPr shadeToTitle="1"><!--before--><a:solidFill>'
        + '<a:srgbClr val="FF0000"/></a:solidFill><x:owned keep="yes"/>'
        + '<a:effectLst><a:outerShdw blurRad="1"/></a:effectLst><!--after-->'
        + '</p:bgPr></p:bg><p:spTree><x:neighbor value="KEEP"/></p:spTree>'
        + '</p:cSld></p:sld>',
    );

    replaceSlideBackground(pkg, SLIDE_URI, { kind: 'none' });
    const xml = slideXml(pkg);
    expect(xml).toContain('<p:bgPr shadeToTitle="1"><!--before--><a:noFill/>');
    expect(xml).toContain('<x:owned keep="yes"/>');
    expect(xml).toContain('<a:effectLst><a:outerShdw blurRad="1"/></a:effectLst><!--after-->');
    expect(xml).toContain('<p:spTree><x:neighbor value="KEEP"/></p:spTree>');
  });

  it('treats repeated normalized assignments and absent clears as exact no-ops', () => {
    const pkg = backgroundPackage('');
    const solid = {
      kind: 'solid',
      color: { kind: 'srgb', value: 'FF3399' },
      transparency: 50,
    } as const;
    replaceSlideBackground(pkg, SLIDE_URI, solid);
    const solidSnapshot = snapshot(pkg);
    replaceSlideBackground(pkg, SLIDE_URI, solid);
    expect(snapshot(pkg)).toEqual(solidSnapshot);

    const gradient = {
      kind: 'linear-gradient',
      angle: 0,
      stops: validStops(),
    } as const;
    replaceSlideBackground(pkg, SLIDE_URI, gradient);
    const gradientSnapshot = snapshot(pkg);
    replaceSlideBackground(pkg, SLIDE_URI, gradient);
    expect(snapshot(pkg)).toEqual(gradientSnapshot);

    replaceSlideBackground(pkg, SLIDE_URI, undefined);
    const clearSnapshot = snapshot(pkg);
    replaceSlideBackground(pkg, SLIDE_URI, undefined);
    expect(snapshot(pkg)).toEqual(clearSnapshot);
  });

  it.each([
    ['background reference', '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'],
    ['pattern fill', '<p:bg><p:bgPr><a:pattFill prst="pct5"/></p:bgPr></p:bg>'],
    ['empty properties', '<p:bg><p:bgPr><a:effectLst/></p:bgPr></p:bg>'],
    [
      'duplicate properties',
      '<p:bg><p:bgPr><a:noFill/></p:bgPr><p:bgPr><a:noFill/></p:bgPr></p:bg>',
    ],
  ] as const)('replaces opaque %s state with one canonical background', (_name, background) => {
    const pkg = backgroundPackage(background);
    replaceSlideBackground(pkg, SLIDE_URI, { kind: 'none' });
    expect(readSlideBackground(pkg, SLIDE_URI)).toEqual({ kind: 'none' });
    expect((slideXml(pkg).match(/<p:bg>/g) ?? [])).toHaveLength(1);
    expect((slideXml(pkg).match(/<p:bgPr>/g) ?? [])).toHaveLength(1);
  });

  it('clears an opaque direct background instead of confusing it with inheritance', () => {
    const pkg = backgroundPackage(
      '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>',
    );
    expect(readSlideBackground(pkg, SLIDE_URI)).toBeUndefined();
    replaceSlideBackground(pkg, SLIDE_URI, undefined);
    expect(slideXml(pkg)).not.toContain('<p:bg>');
  });

  it('collapses duplicate direct backgrounds without touching descendant or wrong-namespace traps', () => {
    const pkg = packageFromSlide(
      `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" `
        + `xmlns:x="urn:opaque"><p:cSld><p:bg><p:bgPr><a:noFill/></p:bgPr></p:bg>`
        + '<x:bg><x:keep/></x:bg><p:bg><p:bgPr><a:solidFill>'
        + '<a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg><p:spTree>'
        + '<p:extLst><p:ext><p:bg><p:bgPr><a:noFill/></p:bgPr></p:bg></p:ext></p:extLst>'
        + '</p:spTree></p:cSld></p:sld>',
    );
    replaceSlideBackground(pkg, SLIDE_URI, {
      kind: 'solid',
      color: { kind: 'srgb', value: '112233' },
    });
    const xml = slideXml(pkg);
    expect((xml.match(/<p:cSld><p:bg>/g) ?? [])).toHaveLength(1);
    expect(xml).toContain('<x:bg><x:keep/></x:bg>');
    expect(xml).toContain(
      '<p:extLst><p:ext><p:bg><p:bgPr><a:noFill/></p:bgPr></p:bg></p:ext></p:extLst>',
    );
  });

  it('rejects invalid values before mutation and rolls back outer transactions', () => {
    const pkg = backgroundPackage('');
    const before = snapshot(pkg);
    expect(() => replaceSlideBackground(pkg, SLIDE_URI, {
      kind: 'solid',
      color: { kind: 'srgb', value: 'FFF' },
    })).toThrow(TypeError);
    expect(snapshot(pkg)).toEqual(before);

    expect(() => pkg.transaction(() => {
      replaceSlideBackground(pkg, SLIDE_URI, { kind: 'none' });
      throw new Error('rollback');
    })).toThrow('rollback');
    expect(snapshot(pkg)).toEqual(before);

    const malformed = packageFromSlide(
      `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}"/>`,
    );
    const malformedBefore = snapshot(malformed);
    expect(() => replaceSlideBackground(malformed, SLIDE_URI, { kind: 'none' }))
      .toThrow(/editable cSld/);
    expect(snapshot(malformed)).toEqual(malformedBefore);
  });
});

describe('raster slide background lifecycle', () => {
  it.each([
    ['image/png', '.png'],
    ['image/jpeg', '.jpeg'],
    ['image/gif', '.gif'],
  ] as const)('creates a detached canonical %s background', async (contentType, extension) => {
    const pkg = backgroundPackage('');
    const bytes = Uint8Array.of(1, 2, 3);
    replaceSlideBackground(pkg, SLIDE_URI, {
      kind: 'image',
      contentType,
      bytes,
    });
    bytes[0] = 9;

    const relationships = pkg.relationships(SLIDE_URI).filter(({ type }) =>
      type === IMAGE_RELATIONSHIP);
    expect(relationships).toHaveLength(1);
    expect(relationships[0]).toMatchObject({ targetMode: 'Internal' });
    const target = relationships[0]!.resolvedTarget!;
    expect(target).toMatch(new RegExp(`/ppt/media/background\\d+\\${extension}$`));
    expect(pkg.requirePart(target)).toMatchObject({ contentType });
    expect(pkg.requirePart(target).bytes).toEqual(Uint8Array.of(1, 2, 3));
    expect(slideXml(pkg)).toContain(
      `<a:blipFill><a:blip r:embed="${relationships[0]!.id}"/>`
        + '<a:stretch><a:fillRect/></a:stretch></a:blipFill>',
    );
    const read = readSlideBackground(pkg, SLIDE_URI);
    expect(read).toEqual({
      kind: 'image',
      contentType,
      bytes: Uint8Array.of(1, 2, 3),
    });
    if (read?.kind !== 'image') throw new Error('Expected image background');
    read.bytes[0] = 8;
    expect(pkg.requirePart(target).bytes).toEqual(Uint8Array.of(1, 2, 3));
    expect(slideBackgroundMediaTargets(pkg, SLIDE_URI)).toEqual([target]);

    const reopened = await OpcPackage.open(await pkg.write());
    expect(readSlideBackground(reopened, SLIDE_URI)).toEqual({
      kind: 'image',
      contentType,
      bytes: Uint8Array.of(1, 2, 3),
    });
  });

  it('no-ops exact image values, updates exclusive targets, and preserves the relationship id across MIME changes', () => {
    const pkg = backgroundPackage('');
    replaceSlideBackground(pkg, SLIDE_URI, {
      kind: 'image',
      contentType: 'image/png',
      bytes: Uint8Array.of(1, 2, 3),
    });
    const relationship = pkg.relationships(SLIDE_URI).find(({ type }) =>
      type === IMAGE_RELATIONSHIP)!;
    const originalTarget = relationship.resolvedTarget!;
    const exactSnapshot = snapshot(pkg);
    replaceSlideBackground(pkg, SLIDE_URI, {
      kind: 'image',
      contentType: 'image/png',
      bytes: Uint8Array.of(1, 2, 3),
    });
    expect(snapshot(pkg)).toEqual(exactSnapshot);

    replaceSlideBackground(pkg, SLIDE_URI, {
      kind: 'image',
      contentType: 'image/png',
      bytes: Uint8Array.of(4, 5, 6),
    });
    expect(pkg.relationships(SLIDE_URI).find(({ id }) => id === relationship.id)?.resolvedTarget)
      .toBe(originalTarget);
    expect(pkg.requirePart(originalTarget).bytes).toEqual(Uint8Array.of(4, 5, 6));

    replaceSlideBackground(pkg, SLIDE_URI, {
      kind: 'image',
      contentType: 'image/jpeg',
      bytes: Uint8Array.of(7, 8),
    });
    const changed = pkg.relationships(SLIDE_URI).find(({ id }) => id === relationship.id)!;
    expect(changed.resolvedTarget).toMatch(/\.jpeg$/);
    expect(changed.resolvedTarget).not.toBe(originalTarget);
    expect(pkg.hasPart(originalTarget)).toBe(false);

    replaceSlideBackground(pkg, SLIDE_URI, {
      kind: 'solid',
      color: { kind: 'srgb', value: '112233' },
    });
    expect(pkg.relationships(SLIDE_URI).some(({ id }) => id === relationship.id)).toBe(false);
    expect(pkg.hasPart(changed.resolvedTarget!)).toBe(false);
  });

  it.each([
    ['inherited background', undefined],
    ['no fill', { kind: 'none' }],
    ['solid fill', { kind: 'solid', color: { kind: 'srgb', value: '112233' } }],
    ['gradient fill', {
      kind: 'linear-gradient',
      angle: 45,
      stops: [
        { offset: 0, color: 'FFFFFF' },
        { offset: 1, color: '000000' },
      ],
    }],
  ] as const)('cleans image resources when replacing with %s', (_name, replacement) => {
    const pkg = backgroundPackage('');
    replaceSlideBackground(pkg, SLIDE_URI, {
      kind: 'solid',
      color: { kind: 'scheme', value: 'accent1' },
    });
    replaceSlideBackground(pkg, SLIDE_URI, {
      kind: 'image',
      contentType: 'image/png',
      bytes: Uint8Array.of(1, 2, 3),
    });
    const relationship = pkg.relationships(SLIDE_URI).find(({ type }) =>
      type === IMAGE_RELATIONSHIP)!;
    const target = relationship.resolvedTarget!;

    replaceSlideBackground(pkg, SLIDE_URI, replacement);

    expect(pkg.relationships(SLIDE_URI).some(({ id }) => id === relationship.id)).toBe(false);
    expect(pkg.hasPart(target)).toBe(false);
    const read = readSlideBackground(pkg, SLIDE_URI);
    if (replacement === undefined) {
      expect(read).toBeUndefined();
    } else {
      expect(read?.kind).toBe(replacement.kind);
    }
  });

  it('isolates a relationship id shared by the background and a picture', () => {
    const oldTarget = '/ppt/media/shared.png';
    const pkg = packageFromSlide(
      `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" `
        + `xmlns:r="${RELATIONSHIP_NAMESPACE}"><p:cSld><p:bg><p:bgPr>`
        + '<a:blipFill><a:blip r:embed="rId7"/></a:blipFill></p:bgPr></p:bg>'
        + '<p:spTree><p:pic><p:blipFill><a:blip r:embed="rId7"/></p:blipFill>'
        + '</p:pic></p:spTree></p:cSld></p:sld>',
      {
        parts: [{
          uri: oldTarget,
          contentType: 'image/png',
          bytes: Uint8Array.of(1),
        }],
        relationships: [{ id: 'rId7', type: IMAGE_RELATIONSHIP, target: oldTarget }],
      },
    );

    replaceSlideBackground(pkg, SLIDE_URI, {
      kind: 'image',
      contentType: 'image/png',
      bytes: Uint8Array.of(2),
    });
    const relationships = pkg.relationships(SLIDE_URI).filter(({ type }) =>
      type === IMAGE_RELATIONSHIP);
    expect(relationships).toHaveLength(2);
    expect(pkg.requirePart(oldTarget).bytes).toEqual(Uint8Array.of(1));
    expect(slideXml(pkg)).toContain('<p:blipFill><a:blip r:embed="rId7"/></p:blipFill>');
    const replacement = relationships.find(({ id }) => id !== 'rId7')!;
    expect(slideXml(pkg)).toContain(`<a:blipFill><a:blip r:embed="${replacement.id}"/>`);
    expect(pkg.requirePart(replacement.resolvedTarget!).bytes).toEqual(Uint8Array.of(2));

    replaceSlideBackground(pkg, SLIDE_URI, undefined);
    expect(pkg.hasPart(replacement.resolvedTarget!)).toBe(false);
    expect(pkg.hasPart(oldTarget)).toBe(true);
    expect(pkg.relationships(SLIDE_URI).map(({ id }) => id)).toEqual(['rId7']);
    expect(slideXml(pkg)).toContain('<p:blipFill><a:blip r:embed="rId7"/></p:blipFill>');
  });

  it('isolates a media target shared by separate background and picture relationships', () => {
    const oldTarget = '/ppt/media/shared.png';
    const pkg = packageFromSlide(
      `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" `
        + `xmlns:r="${RELATIONSHIP_NAMESPACE}"><p:cSld><p:bg><p:bgPr>`
        + '<a:blipFill><a:blip r:embed="rId7"/></a:blipFill></p:bgPr></p:bg>'
        + '<p:spTree><p:pic><p:blipFill><a:blip r:embed="rId8"/></p:blipFill>'
        + '</p:pic></p:spTree></p:cSld></p:sld>',
      {
        parts: [{
          uri: oldTarget,
          contentType: 'image/png',
          bytes: Uint8Array.of(1),
        }],
        relationships: [
          { id: 'rId7', type: IMAGE_RELATIONSHIP, target: oldTarget },
          { id: 'rId8', type: IMAGE_RELATIONSHIP, target: oldTarget },
        ],
      },
    );

    replaceSlideBackground(pkg, SLIDE_URI, {
      kind: 'image',
      contentType: 'image/png',
      bytes: Uint8Array.of(2),
    });

    const backgroundRelationship = pkg.relationships(SLIDE_URI).find(({ id }) =>
      id === 'rId7')!;
    const pictureRelationship = pkg.relationships(SLIDE_URI).find(({ id }) =>
      id === 'rId8')!;
    expect(backgroundRelationship.resolvedTarget).not.toBe(oldTarget);
    expect(pictureRelationship.resolvedTarget).toBe(oldTarget);
    expect(pkg.requirePart(oldTarget).bytes).toEqual(Uint8Array.of(1));
    expect(pkg.requirePart(backgroundRelationship.resolvedTarget!).bytes).toEqual(Uint8Array.of(2));
    expect(slideXml(pkg)).toContain('<a:blip r:embed="rId7"/>');
    expect(slideXml(pkg)).toContain('<a:blip r:embed="rId8"/>');

    replaceSlideBackground(pkg, SLIDE_URI, undefined);
    expect(pkg.hasPart(backgroundRelationship.resolvedTarget!)).toBe(false);
    expect(pkg.hasPart(oldTarget)).toBe(true);
    expect(pkg.relationships(SLIDE_URI).map(({ id }) => id)).toEqual(['rId8']);
  });

  it('cleans image relationships hidden inside an opaque direct background', () => {
    const target = '/ppt/media/opaque.png';
    const pkg = backgroundPackage(
      '<p:bg><p:bgPr><a:pattFill><a:fgClr><a:schemeClr val="accent1" '
        + 'r:embed="rId7"/></a:fgClr></a:pattFill></p:bgPr></p:bg>',
      {
        parts: [{ uri: target, contentType: 'image/png', bytes: Uint8Array.of(1) }],
        relationships: [{ id: 'rId7', type: IMAGE_RELATIONSHIP, target }],
      },
    );
    expect(readSlideBackground(pkg, SLIDE_URI)).toBeUndefined();

    replaceSlideBackground(pkg, SLIDE_URI, {
      kind: 'solid',
      color: { kind: 'srgb', value: '112233' },
    });

    expect(pkg.relationships(SLIDE_URI)).toEqual([]);
    expect(pkg.hasPart(target)).toBe(false);
    expect(readSlideBackground(pkg, SLIDE_URI)).toEqual({
      kind: 'solid',
      color: { kind: 'srgb', value: '112233' },
    });
  });

  it('cleans only unreferenced background resources and restores all state on rollback', () => {
    const pkg = backgroundPackage('');
    replaceSlideBackground(pkg, SLIDE_URI, {
      kind: 'image',
      contentType: 'image/gif',
      bytes: Uint8Array.of(1, 2),
    });
    const target = slideBackgroundMediaTargets(pkg, SLIDE_URI)[0]!;
    const before = snapshot(pkg);
    expect(() => pkg.transaction(() => {
      replaceSlideBackground(pkg, SLIDE_URI, {
        kind: 'image',
        contentType: 'image/png',
        bytes: Uint8Array.of(9),
      });
      throw new Error('rollback image background');
    })).toThrow('rollback image background');
    expect(snapshot(pkg)).toEqual(before);

    replaceSlideBackground(pkg, SLIDE_URI, undefined);
    expect(pkg.hasPart(target)).toBe(false);
    expect(pkg.relationships(SLIDE_URI).filter(({ type }) => type === IMAGE_RELATIONSHIP))
      .toEqual([]);
    expect(slideBackgroundMediaTargets(pkg, SLIDE_URI)).toEqual([]);
  });
});

function validStops() {
  return [
    { offset: 0, color: 'FFFFFF' },
    { offset: 1, color: '000000' },
  ];
}

interface FixtureOptions {
  readonly parts?: readonly {
    readonly uri: string;
    readonly contentType: string;
    readonly bytes: Uint8Array;
  }[];
  readonly relationships?: readonly {
    readonly id: string;
    readonly type: string;
    readonly target: string;
    readonly targetMode?: 'Internal' | 'External';
  }[];
}

function backgroundPackage(background: string, options: FixtureOptions = {}): OpcPackage {
  return packageFromSlide(
    `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" xmlns:a="${DRAWING_NAMESPACE}" `
      + `xmlns:r="${RELATIONSHIP_NAMESPACE}"><p:cSld>${background}<p:spTree/></p:cSld></p:sld>`,
    options,
  );
}

function ownerBackgroundPackage(
  owner: (typeof BACKGROUND_OWNER_FIXTURES)[number],
  background: string,
  options: FixtureOptions = {},
): OpcPackage {
  return packageFromBackgroundOwner(
    owner,
    `<p:${owner.root} xmlns:p="${PRESENTATION_NAMESPACE}" `
      + `xmlns:a="${DRAWING_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}" `
      + 'xmlns:x="urn:not-presentation"><p:cSld>'
      + `${background}<p:spTree/></p:cSld></p:${owner.root}>`,
    options,
  );
}

function packageFromBackgroundOwner(
  owner: (typeof BACKGROUND_OWNER_FIXTURES)[number],
  xml: string,
  options: FixtureOptions = {},
): OpcPackage {
  const pkg = OpcPackage.create();
  pkg.setPart(owner.uri, xml, owner.contentType);
  for (const part of options.parts ?? []) {
    pkg.setPart(part.uri, part.bytes, part.contentType);
  }
  const relationships = options.relationships ?? [];
  if (relationships.length > 0) {
    const entries = relationships.map((relationship) => {
      const targetMode = relationship.targetMode ?? 'Internal';
      const target = targetMode === 'External'
        ? relationship.target
        : relativeRelationshipTarget(owner.uri, relationship.target);
      return `<Relationship Id="${relationship.id}" Type="${relationship.type}" `
        + `Target="${target}"${targetMode === 'External' ? ' TargetMode="External"' : ''}/>`;
    }).join('');
    const nameStart = owner.uri.lastIndexOf('/') + 1;
    const relationshipsUri = `${owner.uri.slice(0, nameStart)}_rels/`
      + `${owner.uri.slice(nameStart)}.rels`;
    pkg.setPart(
      relationshipsUri,
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + `${entries}</Relationships>`,
      'application/vnd.openxmlformats-package.relationships+xml',
    );
  }
  return pkg;
}

function packageFromSlide(slideXml: string, options: FixtureOptions = {}): OpcPackage {
  const pkg = OpcPackage.create();
  pkg.setPart(SLIDE_URI, slideXml, SLIDE_CONTENT_TYPE);
  for (const part of options.parts ?? []) {
    pkg.setPart(part.uri, part.bytes, part.contentType);
  }
  const relationships = options.relationships ?? [];
  if (relationships.length > 0) {
    const entries = relationships.map((relationship) => {
      const targetMode = relationship.targetMode ?? 'Internal';
      const target = targetMode === 'External'
        ? relationship.target
        : relativeRelationshipTarget(SLIDE_URI, relationship.target);
      return `<Relationship Id="${relationship.id}" Type="${relationship.type}" `
        + `Target="${target}"${targetMode === 'External' ? ' TargetMode="External"' : ''}/>`;
    }).join('');
    pkg.setPart(
      '/ppt/slides/_rels/slide1.xml.rels',
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
        + `${entries}</Relationships>`,
      'application/vnd.openxmlformats-package.relationships+xml',
    );
  }
  return pkg;
}

function snapshot(pkg: OpcPackage, partUri = SLIDE_URI) {
  return {
    parts: pkg.parts.map(({ uri, contentType, bytes }) => ({
      uri,
      contentType,
      bytes: [...bytes],
    })),
    relationships: pkg.relationships(partUri).map((relationship) => ({ ...relationship })),
    graph: pkg.graph.map((node) => ({
      uri: node.uri,
      contentType: node.contentType,
      outgoing: node.outgoing.map((relationship) => ({ ...relationship })),
      incoming: node.incoming.map(({ sourceUri, relationship }) => ({
        sourceUri,
        relationship: { ...relationship },
      })),
    })),
    mutations: pkg.mutations.map((mutation) => ({ ...mutation })),
  };
}

function ownerXml(
  pkg: OpcPackage,
  owner: (typeof BACKGROUND_OWNER_FIXTURES)[number],
): string {
  return new TextDecoder().decode(pkg.requirePart(owner.uri).bytes);
}

function slideXml(pkg: OpcPackage): string {
  return new TextDecoder().decode(pkg.requirePart(SLIDE_URI).bytes);
}
