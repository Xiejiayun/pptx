import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import {
  readPresentationCreatedAt,
  replacePresentationCreatedAt,
} from './presentation-created-at.internal.js';

const CORE_RELATIONSHIP =
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const CORE_CONTENT_TYPE =
  'application/vnd.openxmlformats-package.core-properties+xml';
const RELATIONSHIPS_CONTENT_TYPE =
  'application/vnd.openxmlformats-package.relationships+xml';
const CORE_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const DUBLIN_CORE_NAMESPACE = 'http://purl.org/dc/elements/1.1/';
const DCTERMS_NAMESPACE = 'http://purl.org/dc/terms/';
const XSI_NAMESPACE = 'http://www.w3.org/2001/XMLSchema-instance';

const coreXml = (
  children: string,
  root = 'cp:coreProperties',
  namespaces = `xmlns:cp="${CORE_NAMESPACE}" xmlns:dc="${DUBLIN_CORE_NAMESPACE}" xmlns:dcterms="${DCTERMS_NAMESPACE}" xmlns:xsi="${XSI_NAMESPACE}"`,
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

function replaceCreatedAt(pkg: OpcPackage, value: unknown): void {
  pkg.transaction(() => {
    replacePresentationCreatedAt(pkg, value as never);
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

function partText(pkg: OpcPackage, uri = '/docProps/core.xml'): string {
  return new TextDecoder().decode(pkg.requirePart(uri).bytes);
}

function setRootRelationships(pkg: OpcPackage, relationships: string): void {
  pkg.setPart(
    '/_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`,
    RELATIONSHIPS_CONTENT_TYPE,
  );
}

describe('presentation created-at core property', () => {
  it('reads only one namespace- and type-correct direct simple created timestamp', () => {
    const cases: readonly [OpcPackage, string | undefined][] = [
      [corePackage(), undefined],
      [corePackage(coreXml('<dc:creator>Creator only</dc:creator>')), undefined],
      [corePackage(coreXml(
        '<d:created i:type="d:W3CDTF">2024-02-29T23:59:59.123456+14:00</d:created>',
        'c:coreProperties',
        `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DCTERMS_NAMESPACE}" xmlns:i="${XSI_NAMESPACE}"`,
      ), { uri: '/metadata/core7.xml' }), '2024-02-29T23:59:59.123456+14:00'],
      [corePackage(coreXml(
        `<created xmlns="${DCTERMS_NAMESPACE}" xmlns:i="${XSI_NAMESPACE}" i:type="d:W3CDTF" xmlns:d="${DCTERMS_NAMESPACE}">2026-07-30T00:00:00Z</created>`,
      )), '2026-07-30T00:00:00Z'],
      [corePackage(coreXml(
        '<cp:created>2026-07-30T00:00:00Z</cp:created>',
      )), undefined],
      [corePackage(coreXml(
        '<dcterms:created>2026-07-30T00:00:00Z</dcterms:created>',
      )), undefined],
      [corePackage(coreXml(
        '<dcterms:created xsi:type="dcterms:Other">2026-07-30T00:00:00Z</dcterms:created>',
      )), undefined],
      [corePackage(coreXml(
        '<dcterms:created xsi:type="missing:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>',
      )), undefined],
      [corePackage(coreXml(
        '<dcterms:created xsi:type="W3CDTF">2026-07-30T00:00:00Z</dcterms:created>',
      )), undefined],
      [corePackage(coreXml(
        '<x:opaque xmlns:x="urn:test"><dcterms:created xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created></x:opaque>',
      )), undefined],
      [corePackage(coreXml(
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>'
          + '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-31T00:00:00Z</dcterms:created>',
      )), undefined],
      [corePackage(coreXml(
        `<dcterms:created xmlns:i="${XSI_NAMESPACE}" i:type="dcterms:W3CDTF" xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>`,
      )), undefined],
      [corePackage(coreXml(
        '<dcterms:created xsi:type="dcterms:W3CDTF">Before<x:keep xmlns:x="urn:test"/>After</dcterms:created>',
      )), undefined],
      [corePackage(coreXml(
        '<dcterms:created xsi:type="dcterms:W3CDTF"><![CDATA[2026-07-30T00:00:00Z]]></dcterms:created>',
      )), undefined],
      [corePackage(coreXml(
        '<dcterms:created xsi:type="dcterms:W3CDTF">1900-02-29T00:00:00Z</dcterms:created>',
      )), undefined],
      [corePackage(
        `<x:notCore xmlns:x="urn:test"><dcterms:created xmlns:dcterms="${DCTERMS_NAMESPACE}" xmlns:xsi="${XSI_NAMESPACE}" xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created></x:notCore>`,
      ), undefined],
      [corePackage(coreXml(
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>',
      ), { contentType: 'application/xml' }), undefined],
    ];

    for (const [pkg, expected] of cases) {
      const before = packageSnapshot(pkg);
      expect(readPresentationCreatedAt(pkg)).toBe(expected);
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    const dangling = OpcPackage.create();
    setRootRelationships(
      dangling,
      `<Relationship Id="rId1" Type="${CORE_RELATIONSHIP}" Target="docProps/missing.xml"/>`,
    );
    const danglingBefore = packageSnapshot(dangling);
    expect(readPresentationCreatedAt(dangling)).toBeUndefined();
    expect(packageSnapshot(dangling)).toEqual(danglingBefore);

    const external = OpcPackage.create();
    setRootRelationships(
      external,
      `<Relationship Id="rId1" Type="${CORE_RELATIONSHIP}" Target="https://example.com/core.xml" TargetMode="External"/>`,
    );
    const externalBefore = packageSnapshot(external);
    expect(readPresentationCreatedAt(external)).toBeUndefined();
    expect(packageSnapshot(external)).toEqual(externalBefore);

    const duplicate = corePackage(coreXml(
      '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>',
    ));
    duplicate.transaction(() => {
      duplicate.setPart(
        '/docProps/core2.xml',
        coreXml('<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-31T00:00:00Z</dcterms:created>'),
        CORE_CONTENT_TYPE,
      );
      duplicate.addRelationship('/', {
        id: 'rId2',
        type: CORE_RELATIONSHIP,
        target: 'docProps/core2.xml',
      });
    });
    const duplicateBefore = packageSnapshot(duplicate);
    expect(readPresentationCreatedAt(duplicate)).toBeUndefined();
    expect(packageSnapshot(duplicate)).toEqual(duplicateBefore);
  });

  it('accepts the strict W3CDTF subset and rejects invalid lexical or runtime values', () => {
    const valid = [
      '0001-01-01T00:00:00Z',
      '2000-02-29T23:59:59.0Z',
      '2024-02-29T12:34:56.123456+05:30',
      '2026-07-30T00:00:00-00:30',
      '9999-12-31T23:59:59+14:00',
    ] as const;
    for (const value of valid) {
      const pkg = OpcPackage.create();
      replaceCreatedAt(pkg, value);
      expect(readPresentationCreatedAt(pkg)).toBe(value);
    }

    const invalid = [
      '',
      ' 2026-07-30T00:00:00Z',
      '2026-07-30T00:00:00Z ',
      '0000-01-01T00:00:00Z',
      '10000-01-01T00:00:00Z',
      '1900-02-29T00:00:00Z',
      '2026-02-30T00:00:00Z',
      '2026-00-01T00:00:00Z',
      '2026-13-01T00:00:00Z',
      '2026-07-30',
      '2026-07-30T00:00Z',
      '2026-07-30T00:00:00',
      '2026-07-30T24:00:00Z',
      '2026-07-30T23:60:00Z',
      '2026-07-30T23:59:60Z',
      '2026-07-30T00:00:00.Z',
      '2026-07-30T00:00:00z',
      '2026-07-30T00:00:00+14:01',
      '2026-07-30T00:00:00+15:00',
      '２０２６-07-30T00:00:00Z',
      null,
      true,
      0,
      1n,
      new Date(),
      {},
      [],
      Symbol('createdAt'),
    ] as const;
    for (const value of invalid) {
      const pkg = OpcPackage.create();
      const before = packageSnapshot(pkg);
      expect(() => replaceCreatedAt(pkg, value)).toThrow(TypeError);
      expect(packageSnapshot(pkg)).toEqual(before);
    }
  });

  it('creates only a minimal typed created part without duplicate namespaces', () => {
    const pkg = OpcPackage.create();
    replaceCreatedAt(pkg, '2024-02-29T12:34:56.123456+05:30');

    const relationship = pkg.relationships('/').find(({ type }) => type === CORE_RELATIONSHIP);
    expect(relationship).toMatchObject({
      target: 'docProps/core.xml',
      targetMode: 'Internal',
      resolvedTarget: '/docProps/core.xml',
    });
    expect(pkg.requirePart('/docProps/core.xml').contentType).toBe(CORE_CONTENT_TYPE);
    expect(partText(pkg)).toBe(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="${CORE_NAMESPACE}" xmlns:dcterms="${DCTERMS_NAMESPACE}" xmlns:xsi="${XSI_NAMESPACE}"><dcterms:created xsi:type="dcterms:W3CDTF">2024-02-29T12:34:56.123456+05:30</dcterms:created></cp:coreProperties>`,
    );
    expect(partText(pkg).match(/xmlns:cp=/g)).toHaveLength(1);
    expect(partText(pkg).match(/xmlns:dcterms=/g)).toHaveLength(1);
    expect(partText(pkg).match(/xmlns:xsi=/g)).toHaveLength(1);
    expect(readPresentationCreatedAt(pkg)).toBe('2024-02-29T12:34:56.123456+05:30');

    const occupied = OpcPackage.create();
    const orphan = '<x:orphan xmlns:x="urn:test">KEEP</x:orphan>';
    occupied.setPart('/docProps/core.xml', orphan, 'application/xml');
    replaceCreatedAt(occupied, '2026-07-30T00:00:00Z');
    const allocated = occupied.relationships('/').find(({ type }) => type === CORE_RELATIONSHIP);
    expect(allocated?.resolvedTarget).toBe('/docProps/core1.xml');
    expect(partText(occupied)).toBe(orphan);
    expect(readPresentationCreatedAt(occupied)).toBe('2026-07-30T00:00:00Z');
    expect(partText(occupied, '/docProps/core1.xml')).not.toContain('creator');
    expect(partText(occupied, '/docProps/core1.xml')).not.toContain('modified');
  });

  it('losslessly inserts, replaces, repairs, and clears only created-at', () => {
    const pkg = corePackage(coreXml(
      '\n  <dc:title>Quarterly</dc:title><dc:subject>Forecast</dc:subject>\n'
        + '  <dc:creator>Alice</dc:creator><c:lastModifiedBy>Editor</c:lastModifiedBy><c:revision>7</c:revision>\n'
        + '  <d:modified i:type="d:W3CDTF">2026-07-30T01:00:00Z</d:modified>\n'
        + '  <!--KEEP--><x:opaque xmlns:x="urn:test">KEEP</x:opaque>\n',
      'c:coreProperties',
      `xmlns:c="${CORE_NAMESPACE}" xmlns:dc="${DUBLIN_CORE_NAMESPACE}" xmlns:d="${DCTERMS_NAMESPACE}" xmlns:i="${XSI_NAMESPACE}"`,
    ));
    pkg.setPart('/custom/keep.bin', new Uint8Array([1, 2, 3]), 'application/octet-stream');
    const unrelatedBefore = [...pkg.requirePart('/custom/keep.bin').bytes];

    replaceCreatedAt(pkg, '2024-02-29T12:34:56.123+05:30');
    let xml = partText(pkg);
    expect(xml).toContain(
      '<d:created i:type="d:W3CDTF">2024-02-29T12:34:56.123+05:30</d:created>',
    );
    expect(xml).toContain('<d:modified i:type="d:W3CDTF">2026-07-30T01:00:00Z</d:modified>');
    expect(xml).toContain('<dc:creator>Alice</dc:creator>');
    expect(xml).toContain('<c:lastModifiedBy>Editor</c:lastModifiedBy>');
    expect(xml).toContain('<c:revision>7</c:revision>');
    expect(xml).toContain('<!--KEEP-->');
    expect(xml).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    expect([...pkg.requirePart('/custom/keep.bin').bytes]).toEqual(unrelatedBefore);

    const beforeSame = packageSnapshot(pkg);
    replaceCreatedAt(pkg, '2024-02-29T12:34:56.123+05:30');
    expect(packageSnapshot(pkg)).toEqual(beforeSame);

    replaceCreatedAt(pkg, '2026-07-30T00:00:00Z');
    expect(readPresentationCreatedAt(pkg)).toBe('2026-07-30T00:00:00Z');
    expect(partText(pkg)).toContain(
      '<d:created i:type="d:W3CDTF">2026-07-30T00:00:00Z</d:created>',
    );

    replaceCreatedAt(pkg, undefined);
    expect(readPresentationCreatedAt(pkg)).toBeUndefined();
    xml = partText(pkg);
    expect(xml).not.toContain('<d:created');
    expect(xml).toContain('<d:modified i:type="d:W3CDTF">2026-07-30T01:00:00Z</d:modified>');
    expect(xml).toContain('<dc:creator>Alice</dc:creator>');
    expect(pkg.relationships('/').filter(({ type }) => type === CORE_RELATIONSHIP)).toHaveLength(1);

    const beforeAbsentClear = packageSnapshot(pkg);
    replaceCreatedAt(pkg, undefined);
    expect(packageSnapshot(pkg)).toEqual(beforeAbsentClear);
  });

  it('repairs type state and safely expands self-closing state', () => {
    const numericEntity = corePackage(coreXml(
      '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-30T00&#58;00&#58;00Z</dcterms:created>',
    ));
    const beforeEntityNoOp = packageSnapshot(numericEntity);
    replaceCreatedAt(numericEntity, '2026-07-30T00:00:00Z');
    expect(packageSnapshot(numericEntity)).toEqual(beforeEntityNoOp);

    const missingType = corePackage(coreXml(
      '<d:created custom="KEEP">2026-07-30T00:00:00Z</d:created>',
      'c:coreProperties',
      `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DCTERMS_NAMESPACE}" xmlns:i="${XSI_NAMESPACE}"`,
    ));
    replaceCreatedAt(missingType, '2026-07-30T00:00:00Z');
    expect(partText(missingType)).toContain(
      '<d:created custom="KEEP" i:type="d:W3CDTF">2026-07-30T00:00:00Z</d:created>',
    );

    const wrongType = corePackage(coreXml(
      '<d:created i:type="d:Other">invalid</d:created>',
      'c:coreProperties',
      `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DCTERMS_NAMESPACE}" xmlns:i="${XSI_NAMESPACE}"`,
    ));
    replaceCreatedAt(wrongType, '2024-02-29T12:34:56Z');
    expect(partText(wrongType)).toContain(
      '<d:created i:type="d:W3CDTF">2024-02-29T12:34:56Z</d:created>',
    );

    const wrongNamespaceType = corePackage(coreXml(
      '<d:created x:type="KEEP" xmlns:x="urn:wrong">2026-07-30T00:00:00Z</d:created>',
      'c:coreProperties',
      `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DCTERMS_NAMESPACE}"`,
    ));
    replaceCreatedAt(wrongNamespaceType, '2026-07-30T00:00:00Z');
    expect(partText(wrongNamespaceType)).toContain('x:type="KEEP"');
    expect(partText(wrongNamespaceType)).toContain(`xmlns:xsi="${XSI_NAMESPACE}"`);
    expect(partText(wrongNamespaceType)).toContain('xsi:type="d:W3CDTF"');

    const propertyLocal = corePackage(coreXml(
      `<d:created xmlns:d="${DCTERMS_NAMESPACE}" xmlns:i="${XSI_NAMESPACE}" i:type="d:Other">2026-07-30T00:00:00Z</d:created>`,
      'c:coreProperties',
      `xmlns:c="${CORE_NAMESPACE}"`,
    ));
    replaceCreatedAt(propertyLocal, '2026-07-30T00:00:00Z');
    expect(readPresentationCreatedAt(propertyLocal)).toBe('2026-07-30T00:00:00Z');
    expect(partText(propertyLocal)).toContain('i:type="d:W3CDTF"');

    const selfClosingRoot = corePackage(
      `<?xml version="1.0"?><c:coreProperties xmlns:c="${CORE_NAMESPACE}"/>`,
    );
    replaceCreatedAt(selfClosingRoot, '2026-07-30T00:00:00Z');
    expect(readPresentationCreatedAt(selfClosingRoot)).toBe('2026-07-30T00:00:00Z');
    expect(partText(selfClosingRoot)).toContain(
      `<dcterms:created xmlns:dcterms="${DCTERMS_NAMESPACE}" xmlns:xsi="${XSI_NAMESPACE}" xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>`,
    );

    const selfClosingProperty = corePackage(coreXml(
      '<d:created custom="KEEP"/>',
      'c:coreProperties',
      `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DCTERMS_NAMESPACE}" xmlns:i="${XSI_NAMESPACE}"`,
    ));
    replaceCreatedAt(selfClosingProperty, '2026-07-30T00:00:00Z');
    expect(partText(selfClosingProperty)).toContain(
      '<d:created custom="KEEP" i:type="d:W3CDTF">2026-07-30T00:00:00Z</d:created>',
    );
  });

  it('rejects unsafe ownership without mutation', () => {
    const unsafe: OpcPackage[] = [
      corePackage(coreXml(
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>'
          + '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-31T00:00:00Z</dcterms:created>',
      )),
      corePackage(coreXml(
        `<dcterms:created xmlns:i="${XSI_NAMESPACE}" i:type="dcterms:W3CDTF" xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>`,
      )),
      corePackage(coreXml(
        '<dcterms:created xsi:type="dcterms:W3CDTF">Before<x:keep xmlns:x="urn:test"/>After</dcterms:created>',
      )),
      corePackage(coreXml(
        '<dcterms:created xsi:type="dcterms:W3CDTF"><![CDATA[2026-07-30T00:00:00Z]]></dcterms:created>',
      )),
      corePackage('<x:notCore xmlns:x="urn:test"/>'),
      corePackage(coreXml(
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>',
      ), { contentType: 'application/xml' }),
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

    const duplicate = corePackage(coreXml(
      '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>',
    ));
    duplicate.transaction(() => {
      duplicate.setPart(
        '/docProps/core2.xml',
        coreXml('<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-31T00:00:00Z</dcterms:created>'),
        CORE_CONTENT_TYPE,
      );
      duplicate.addRelationship('/', {
        id: 'rId2',
        type: CORE_RELATIONSHIP,
        target: 'docProps/core2.xml',
      });
    });
    unsafe.push(duplicate);

    for (const pkg of unsafe) {
      const before = packageSnapshot(pkg);
      expect(() => replaceCreatedAt(pkg, '2026-07-30T00:00:00Z')).toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
      expect(() => replaceCreatedAt(pkg, undefined)).toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
    }
  });

  it('rolls back minimal creation and existing created-at edits', () => {
    const created = OpcPackage.create();
    const createdBefore = packageSnapshot(created);

    expect(() => created.transaction(() => {
      replacePresentationCreatedAt(created, '2026-07-30T00:00:00Z');
      expect(readPresentationCreatedAt(created)).toBe('2026-07-30T00:00:00Z');
      throw new Error('restore created-at creation');
    })).toThrow('restore created-at creation');
    expect(packageSnapshot(created)).toEqual(createdBefore);
    expect(created.hasPart('/docProps/core.xml')).toBe(false);
    expect(created.relationships('/')).toHaveLength(0);

    const existing = corePackage(coreXml(
      '<dcterms:created xsi:type="dcterms:W3CDTF">2024-02-29T12:34:56Z</dcterms:created>'
        + '<dcterms:modified xsi:type="dcterms:W3CDTF">2024-02-29T13:34:56Z</dcterms:modified>',
    ));
    const existingBefore = packageSnapshot(existing);
    expect(() => existing.transaction(() => {
      replacePresentationCreatedAt(existing, '2026-07-30T00:00:00Z');
      expect(readPresentationCreatedAt(existing)).toBe('2026-07-30T00:00:00Z');
      throw new Error('restore created-at edit');
    })).toThrow('restore created-at edit');
    expect(packageSnapshot(existing)).toEqual(existingBefore);
    expect(readPresentationCreatedAt(existing)).toBe('2024-02-29T12:34:56Z');
    expect(partText(existing)).toContain(
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2024-02-29T13:34:56Z</dcterms:modified>',
    );
  });
});
