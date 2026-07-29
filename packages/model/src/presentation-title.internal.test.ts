import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import {
  readPresentationTitle,
  replacePresentationTitle,
} from './presentation-title.internal.js';

const CORE_RELATIONSHIP =
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const CORE_CONTENT_TYPE =
  'application/vnd.openxmlformats-package.core-properties+xml';
const RELATIONSHIPS_CONTENT_TYPE =
  'application/vnd.openxmlformats-package.relationships+xml';
const CORE_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const DUBLIN_CORE_NAMESPACE = 'http://purl.org/dc/elements/1.1/';

const coreXml = (
  children: string,
  root = 'cp:coreProperties',
  namespaces = `xmlns:cp="${CORE_NAMESPACE}" xmlns:dc="${DUBLIN_CORE_NAMESPACE}"`,
): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><${root} ${namespaces}>${children}</${root}>`;

function corePackage(
  xml?: string,
  options: { readonly uri?: string; readonly contentType?: string } = {},
): OpcPackage {
  const pkg = OpcPackage.create();
  if (xml !== undefined) {
    const uri = options.uri ?? '/docProps/core.xml';
    pkg.transaction(() => {
      pkg.setPart(uri, xml, options.contentType ?? CORE_CONTENT_TYPE);
      pkg.addRelationship('/', {
        id: 'rId1',
        type: CORE_RELATIONSHIP,
        target: uri.slice(1),
      });
    });
  }
  return pkg;
}

function replaceTitle(pkg: OpcPackage, value: unknown): void {
  pkg.transaction(() => {
    replacePresentationTitle(pkg, value as never);
  });
}

function packageSnapshot(pkg: OpcPackage): {
  readonly parts: readonly (readonly [string, string, readonly number[]])[];
  readonly relationships: ReturnType<OpcPackage['relationships']>;
  readonly mutations: readonly object[];
} {
  return {
    parts: pkg.parts
      .map(({ uri, contentType, bytes }) => [uri, contentType, [...bytes]] as const)
      .sort(([left], [right]) => left.localeCompare(right)),
    relationships: pkg.relationships('/').map((relationship) => ({ ...relationship })),
    mutations: pkg.mutations.map((mutation) => ({ ...mutation })),
  };
}

function setRootRelationships(pkg: OpcPackage, relationships: string): void {
  pkg.setPart(
    '/_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`,
    RELATIONSHIPS_CONTENT_TYPE,
  );
}

describe('presentation title core property', () => {
  it('reads only one namespace-correct direct simple title without mutation', () => {
    const cases: readonly [OpcPackage, string | undefined][] = [
      [corePackage(), undefined],
      [corePackage(coreXml('<dc:title>Quarterly &amp; Review</dc:title>')), 'Quarterly & Review'],
      [corePackage(
        coreXml(
          '<d:title/>',
          'c:coreProperties',
          `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DUBLIN_CORE_NAMESPACE}"`,
        ),
      ), ''],
      [corePackage(coreXml('<x:title xmlns:x="urn:wrong">Wrong</x:title>')), undefined],
      [corePackage(coreXml('<x:opaque xmlns:x="urn:test"><dc:title>Nested</dc:title></x:opaque>')), undefined],
      [corePackage(coreXml('<dc:title>One</dc:title><dc:title>Two</dc:title>')), undefined],
      [corePackage(coreXml('<dc:title>Before<x:keep xmlns:x="urn:test"/>After</dc:title>')), undefined],
      [corePackage(coreXml('<dc:title><![CDATA[Raw]]></dc:title>')), undefined],
      [corePackage('<x:notCore xmlns:x="urn:test"><dc:title xmlns:dc="http://purl.org/dc/elements/1.1/">Wrong root</dc:title></x:notCore>'), undefined],
      [corePackage(coreXml('<dc:title>Wrong type</dc:title>'), { contentType: 'application/xml' }), undefined],
    ];

    for (const [pkg, expected] of cases) {
      const before = packageSnapshot(pkg);
      expect(readPresentationTitle(pkg)).toBe(expected);
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    const dangling = OpcPackage.create();
    setRootRelationships(
      dangling,
      `<Relationship Id="rId1" Type="${CORE_RELATIONSHIP}" Target="docProps/missing.xml"/>`,
    );
    const danglingBefore = packageSnapshot(dangling);
    expect(readPresentationTitle(dangling)).toBeUndefined();
    expect(packageSnapshot(dangling)).toEqual(danglingBefore);

    const external = OpcPackage.create();
    setRootRelationships(
      external,
      `<Relationship Id="rId1" Type="${CORE_RELATIONSHIP}" Target="https://example.com/core.xml" TargetMode="External"/>`,
    );
    const externalBefore = packageSnapshot(external);
    expect(readPresentationTitle(external)).toBeUndefined();
    expect(packageSnapshot(external)).toEqual(externalBefore);

    const duplicate = corePackage(coreXml('<dc:title>One</dc:title>'));
    duplicate.transaction(() => {
      duplicate.setPart('/docProps/core2.xml', coreXml('<dc:title>Two</dc:title>'), CORE_CONTENT_TYPE);
      duplicate.addRelationship('/', {
        id: 'rId2',
        type: CORE_RELATIONSHIP,
        target: 'docProps/core2.xml',
      });
    });
    const duplicateBefore = packageSnapshot(duplicate);
    expect(readPresentationTitle(duplicate)).toBeUndefined();
    expect(packageSnapshot(duplicate)).toEqual(duplicateBefore);
  });

  it('creates a minimal core part and avoids an occupied canonical URI', () => {
    const pkg = OpcPackage.create();
    replaceTitle(pkg, 'Quarterly & <Review>');

    const relationship = pkg.relationships('/').find(({ type }) => type === CORE_RELATIONSHIP);
    expect(relationship).toMatchObject({
      target: 'docProps/core.xml',
      targetMode: 'Internal',
      resolvedTarget: '/docProps/core.xml',
    });
    expect(pkg.requirePart('/docProps/core.xml').contentType).toBe(CORE_CONTENT_TYPE);
    expect(new TextDecoder().decode(pkg.requirePart('/docProps/core.xml').bytes)).toBe(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="${CORE_NAMESPACE}" xmlns:dc="${DUBLIN_CORE_NAMESPACE}"><dc:title>Quarterly &amp; &lt;Review&gt;</dc:title></cp:coreProperties>`,
    );
    expect(readPresentationTitle(pkg)).toBe('Quarterly & <Review>');

    const occupied = OpcPackage.create();
    const orphan = '<x:orphan xmlns:x="urn:test">KEEP</x:orphan>';
    occupied.setPart('/docProps/core.xml', orphan, 'application/xml');
    replaceTitle(occupied, 'Allocated');
    const allocated = occupied.relationships('/').find(({ type }) => type === CORE_RELATIONSHIP);
    expect(allocated?.resolvedTarget).toBe('/docProps/core1.xml');
    expect(new TextDecoder().decode(occupied.requirePart('/docProps/core.xml').bytes)).toBe(orphan);
    expect(readPresentationTitle(occupied)).toBe('Allocated');
  });

  it('losslessly inserts, replaces, clears, and preserves semantic no-ops', () => {
    const pkg = corePackage(coreXml(
      '<cp:revision>7</cp:revision><x:opaque xmlns:x="urn:test">KEEP</x:opaque>',
      'c:coreProperties',
      `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DUBLIN_CORE_NAMESPACE}" xmlns:cp="${CORE_NAMESPACE}"`,
    ));
    replaceTitle(pkg, 'Inserted & <safe>');
    let xml = new TextDecoder().decode(pkg.requirePart('/docProps/core.xml').bytes);
    expect(xml).toContain('<d:title>Inserted &amp; &lt;safe&gt;</d:title>');
    expect(xml).toContain('<cp:revision>7</cp:revision>');
    expect(xml).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');

    const beforeSame = packageSnapshot(pkg);
    replaceTitle(pkg, 'Inserted & <safe>');
    expect(packageSnapshot(pkg)).toEqual(beforeSame);

    replaceTitle(pkg, '');
    expect(readPresentationTitle(pkg)).toBe('');
    xml = new TextDecoder().decode(pkg.requirePart('/docProps/core.xml').bytes);
    expect(xml).toContain('<d:title></d:title>');

    replaceTitle(pkg, undefined);
    expect(readPresentationTitle(pkg)).toBeUndefined();
    xml = new TextDecoder().decode(pkg.requirePart('/docProps/core.xml').bytes);
    expect(xml).not.toContain('<d:title');
    expect(xml).toContain('<cp:revision>7</cp:revision>');
    expect(pkg.relationships('/').filter(({ type }) => type === CORE_RELATIONSHIP)).toHaveLength(1);

    const beforeAbsentClear = packageSnapshot(pkg);
    replaceTitle(pkg, undefined);
    expect(packageSnapshot(pkg)).toEqual(beforeAbsentClear);

    const numericEntity = corePackage(coreXml('<dc:title>A&#38;B</dc:title>'));
    const beforeEntityNoOp = packageSnapshot(numericEntity);
    replaceTitle(numericEntity, 'A&B');
    expect(packageSnapshot(numericEntity)).toEqual(beforeEntityNoOp);

    const selfClosingRoot = corePackage(
      `<?xml version="1.0"?><c:coreProperties xmlns:c="${CORE_NAMESPACE}"/>`,
    );
    replaceTitle(selfClosingRoot, 'Root expanded');
    expect(readPresentationTitle(selfClosingRoot)).toBe('Root expanded');
    expect(new TextDecoder().decode(selfClosingRoot.requirePart('/docProps/core.xml').bytes))
      .toContain(`<dc:title xmlns:dc="${DUBLIN_CORE_NAMESPACE}">Root expanded</dc:title>`);

    const selfClosingTitle = corePackage(coreXml('<dc:title custom="KEEP"/>'));
    replaceTitle(selfClosingTitle, 'Expanded');
    expect(new TextDecoder().decode(selfClosingTitle.requirePart('/docProps/core.xml').bytes))
      .toContain('<dc:title custom="KEEP">Expanded</dc:title>');
  });

  it('rejects invalid values and unsafe ownership without mutation', () => {
    for (const value of [null, false, 0, {}, [], Symbol('title'), 'bad\u0001title']) {
      const pkg = OpcPackage.create();
      const before = packageSnapshot(pkg);
      expect(() => replaceTitle(pkg, value)).toThrow(TypeError);
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    const unsafe: OpcPackage[] = [
      corePackage(coreXml('<dc:title>One</dc:title><dc:title>Two</dc:title>')),
      corePackage(coreXml('<dc:title>Before<x:keep xmlns:x="urn:test"/>After</dc:title>')),
      corePackage('<x:notCore xmlns:x="urn:test"/>'),
      corePackage(coreXml('<dc:title>Wrong type</dc:title>'), { contentType: 'application/xml' }),
    ];

    const dangling = OpcPackage.create();
    setRootRelationships(
      dangling,
      `<Relationship Id="rId1" Type="${CORE_RELATIONSHIP}" Target="docProps/missing.xml"/>`,
    );
    unsafe.push(dangling);

    const external = OpcPackage.create();
    setRootRelationships(
      external,
      `<Relationship Id="rId1" Type="${CORE_RELATIONSHIP}" Target="https://example.com/core.xml" TargetMode="External"/>`,
    );
    unsafe.push(external);

    const duplicate = corePackage(coreXml('<dc:title>One</dc:title>'));
    duplicate.transaction(() => {
      duplicate.setPart('/docProps/core2.xml', coreXml('<dc:title>Two</dc:title>'), CORE_CONTENT_TYPE);
      duplicate.addRelationship('/', {
        id: 'rId2',
        type: CORE_RELATIONSHIP,
        target: 'docProps/core2.xml',
      });
    });
    unsafe.push(duplicate);

    for (const pkg of unsafe) {
      const before = packageSnapshot(pkg);
      expect(() => replaceTitle(pkg, 'Replacement')).toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
    }
  });

  it('rolls back part, relationship, content types, and journal creation', () => {
    const pkg = OpcPackage.create();
    const before = packageSnapshot(pkg);

    expect(() => pkg.transaction(() => {
      replacePresentationTitle(pkg, 'Temporary');
      expect(readPresentationTitle(pkg)).toBe('Temporary');
      throw new Error('restore title lifecycle');
    })).toThrow('restore title lifecycle');

    expect(packageSnapshot(pkg)).toEqual(before);
    expect(pkg.hasPart('/docProps/core.xml')).toBe(false);
    expect(pkg.relationships('/')).toHaveLength(0);
  });
});
