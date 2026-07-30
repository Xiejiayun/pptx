import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { readPresentationCreatedAt } from './presentation-created-at.internal.js';
import {
  readPresentationModifiedAt,
  replacePresentationModifiedAt,
} from './presentation-modified-at.internal.js';

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

function replaceModifiedAt(pkg: OpcPackage, value: unknown): void {
  pkg.transaction(() => {
    replacePresentationModifiedAt(pkg, value as never);
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

describe('presentation modified-at core property', () => {
  it('reads only one namespace- and type-correct direct simple modified timestamp', () => {
    const cases: readonly [OpcPackage, string | undefined][] = [
      [corePackage(), undefined],
      [corePackage(coreXml('<dc:creator>Creator only</dc:creator>')), undefined],
      [corePackage(coreXml(
        '<d:modified i:type="d:W3CDTF">2024-02-29T23:59:59.123456+14:00</d:modified>',
        'c:coreProperties',
        `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DCTERMS_NAMESPACE}" xmlns:i="${XSI_NAMESPACE}"`,
      ), { uri: '/metadata/core7.xml' }), '2024-02-29T23:59:59.123456+14:00'],
      [corePackage(coreXml(
        `<modified xmlns="${DCTERMS_NAMESPACE}" xmlns:i="${XSI_NAMESPACE}" i:type="d:W3CDTF" xmlns:d="${DCTERMS_NAMESPACE}">2026-07-30T01:02:03Z</modified>`,
      )), '2026-07-30T01:02:03Z'],
      [corePackage(coreXml(
        '<cp:modified>2026-07-30T01:02:03Z</cp:modified>',
      )), undefined],
      [corePackage(coreXml(
        '<dcterms:modified>2026-07-30T01:02:03Z</dcterms:modified>',
      )), undefined],
      [corePackage(coreXml(
        '<dcterms:modified xsi:type="dcterms:Other">2026-07-30T01:02:03Z</dcterms:modified>',
      )), undefined],
      [corePackage(coreXml(
        '<dcterms:modified xsi:type="missing:W3CDTF">2026-07-30T01:02:03Z</dcterms:modified>',
      )), undefined],
      [corePackage(coreXml(
        '<dcterms:modified xsi:type="W3CDTF">2026-07-30T01:02:03Z</dcterms:modified>',
      )), undefined],
      [corePackage(coreXml(
        '<x:opaque xmlns:x="urn:test"><dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-30T01:02:03Z</dcterms:modified></x:opaque>',
      )), undefined],
      [corePackage(coreXml(
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-30T01:02:03Z</dcterms:modified>'
          + '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-31T01:02:03Z</dcterms:modified>',
      )), undefined],
      [corePackage(coreXml(
        `<dcterms:modified xmlns:i="${XSI_NAMESPACE}" i:type="dcterms:W3CDTF" xsi:type="dcterms:W3CDTF">2026-07-30T01:02:03Z</dcterms:modified>`,
      )), undefined],
      [corePackage(coreXml(
        '<dcterms:modified xsi:type="dcterms:W3CDTF">Before<x:keep xmlns:x="urn:test"/>After</dcterms:modified>',
      )), undefined],
      [corePackage(coreXml(
        '<dcterms:modified xsi:type="dcterms:W3CDTF"><![CDATA[2026-07-30T01:02:03Z]]></dcterms:modified>',
      )), undefined],
      [corePackage(coreXml(
        '<dcterms:modified xsi:type="dcterms:W3CDTF">1900-02-29T01:02:03Z</dcterms:modified>',
      )), undefined],
      [corePackage(
        `<x:notCore xmlns:x="urn:test"><dcterms:modified xmlns:dcterms="${DCTERMS_NAMESPACE}" xmlns:xsi="${XSI_NAMESPACE}" xsi:type="dcterms:W3CDTF">2026-07-30T01:02:03Z</dcterms:modified></x:notCore>`,
      ), undefined],
      [corePackage(coreXml(
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-30T01:02:03Z</dcterms:modified>',
      ), { contentType: 'application/xml' }), undefined],
    ];

    for (const [pkg, expected] of cases) {
      const before = packageSnapshot(pkg);
      expect(readPresentationModifiedAt(pkg)).toBe(expected);
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    for (const relationshipXml of [
      `<Relationship Id="rId1" Type="${CORE_RELATIONSHIP}" Target="docProps/missing.xml"/>`,
      `<Relationship Id="rId1" Type="${CORE_RELATIONSHIP}" Target="https://example.com/core.xml" TargetMode="External"/>`,
    ]) {
      const pkg = OpcPackage.create();
      setRootRelationships(pkg, relationshipXml);
      const before = packageSnapshot(pkg);
      expect(readPresentationModifiedAt(pkg)).toBeUndefined();
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    const duplicate = corePackage(coreXml(
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-30T01:02:03Z</dcterms:modified>',
    ));
    duplicate.transaction(() => {
      duplicate.setPart(
        '/docProps/core2.xml',
        coreXml('<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-31T01:02:03Z</dcterms:modified>'),
        CORE_CONTENT_TYPE,
      );
      duplicate.addRelationship('/', {
        id: 'rId2',
        type: CORE_RELATIONSHIP,
        target: 'docProps/core2.xml',
      });
    });
    const duplicateBefore = packageSnapshot(duplicate);
    expect(readPresentationModifiedAt(duplicate)).toBeUndefined();
    expect(packageSnapshot(duplicate)).toEqual(duplicateBefore);
  });

  it('accepts the shared strict W3CDTF subset and rejects invalid values', () => {
    const valid = [
      '0001-01-01T00:00:00Z',
      '2000-02-29T23:59:59.0Z',
      '2024-02-29T12:34:56.123456+05:30',
      '2026-07-30T00:00:00-00:30',
      '9999-12-31T23:59:59+14:00',
    ] as const;
    for (const value of valid) {
      const pkg = OpcPackage.create();
      replaceModifiedAt(pkg, value);
      expect(readPresentationModifiedAt(pkg)).toBe(value);
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
      Symbol('modifiedAt'),
    ] as const;
    for (const value of invalid) {
      const pkg = OpcPackage.create();
      const before = packageSnapshot(pkg);
      expect(() => replaceModifiedAt(pkg, value)).toThrow(TypeError);
      expect(packageSnapshot(pkg)).toEqual(before);
    }
  });

  it('creates a minimal typed modified part and avoids occupied canonical URI', () => {
    const pkg = OpcPackage.create();
    replaceModifiedAt(pkg, '2024-02-29T12:34:56.123456+05:30');

    expect(pkg.relationships('/').find(({ type }) => type === CORE_RELATIONSHIP))
      .toMatchObject({
        target: 'docProps/core.xml',
        targetMode: 'Internal',
        resolvedTarget: '/docProps/core.xml',
      });
    expect(pkg.requirePart('/docProps/core.xml').contentType).toBe(CORE_CONTENT_TYPE);
    expect(partText(pkg)).toBe(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="${CORE_NAMESPACE}" xmlns:dcterms="${DCTERMS_NAMESPACE}" xmlns:xsi="${XSI_NAMESPACE}"><dcterms:modified xsi:type="dcterms:W3CDTF">2024-02-29T12:34:56.123456+05:30</dcterms:modified></cp:coreProperties>`,
    );
    expect(partText(pkg).match(/xmlns:cp=/g)).toHaveLength(1);
    expect(partText(pkg).match(/xmlns:dcterms=/g)).toHaveLength(1);
    expect(partText(pkg).match(/xmlns:xsi=/g)).toHaveLength(1);
    expect(partText(pkg)).not.toContain('created');

    const occupied = OpcPackage.create();
    const orphan = '<x:orphan xmlns:x="urn:test">KEEP</x:orphan>';
    occupied.setPart('/docProps/core.xml', orphan, 'application/xml');
    replaceModifiedAt(occupied, '2026-07-30T01:02:03Z');
    const allocated = occupied.relationships('/').find(({ type }) => type === CORE_RELATIONSHIP);
    expect(allocated?.resolvedTarget).toBe('/docProps/core1.xml');
    expect(partText(occupied)).toBe(orphan);
    expect(readPresentationModifiedAt(occupied)).toBe('2026-07-30T01:02:03Z');
    expect(partText(occupied, '/docProps/core1.xml')).not.toContain('creator');
    expect(partText(occupied, '/docProps/core1.xml')).not.toContain('created');
  });

  it('losslessly inserts, repairs, replaces, and clears only modified-at', () => {
    const pkg = corePackage(coreXml(
      '\n  <dc:title>Quarterly</dc:title><dc:subject>Forecast</dc:subject>\n'
        + '  <dc:creator>Alice</dc:creator><c:lastModifiedBy>Editor</c:lastModifiedBy><c:revision>7</c:revision>\n'
        + '  <d:created i:type="d:W3CDTF">2026-07-29T00:00:00Z</d:created>\n'
        + '  <!--KEEP--><x:opaque xmlns:x="urn:test">KEEP</x:opaque>\n',
      'c:coreProperties',
      `xmlns:c="${CORE_NAMESPACE}" xmlns:dc="${DUBLIN_CORE_NAMESPACE}" xmlns:d="${DCTERMS_NAMESPACE}" xmlns:i="${XSI_NAMESPACE}"`,
    ));
    pkg.setPart('/custom/keep.bin', new Uint8Array([1, 2, 3]), 'application/octet-stream');
    const unrelatedBefore = [...pkg.requirePart('/custom/keep.bin').bytes];

    replaceModifiedAt(pkg, '2024-03-01T01:02:03.456+08:00');
    let xml = partText(pkg);
    expect(xml).toContain(
      '<d:modified i:type="d:W3CDTF">2024-03-01T01:02:03.456+08:00</d:modified>',
    );
    expect(readPresentationCreatedAt(pkg)).toBe('2026-07-29T00:00:00Z');
    expect(xml).toContain('<dc:creator>Alice</dc:creator>');
    expect(xml).toContain('<c:lastModifiedBy>Editor</c:lastModifiedBy>');
    expect(xml).toContain('<c:revision>7</c:revision>');
    expect(xml).toContain('<!--KEEP-->');
    expect(xml).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    expect([...pkg.requirePart('/custom/keep.bin').bytes]).toEqual(unrelatedBefore);

    const beforeSame = packageSnapshot(pkg);
    replaceModifiedAt(pkg, '2024-03-01T01:02:03.456+08:00');
    expect(packageSnapshot(pkg)).toEqual(beforeSame);

    replaceModifiedAt(pkg, '2026-07-30T01:02:03Z');
    expect(readPresentationModifiedAt(pkg)).toBe('2026-07-30T01:02:03Z');
    expect(readPresentationCreatedAt(pkg)).toBe('2026-07-29T00:00:00Z');

    replaceModifiedAt(pkg, undefined);
    expect(readPresentationModifiedAt(pkg)).toBeUndefined();
    xml = partText(pkg);
    expect(xml).not.toContain('<d:modified');
    expect(xml).toContain(
      '<d:created i:type="d:W3CDTF">2026-07-29T00:00:00Z</d:created>',
    );
    expect(xml).toContain('<c:lastModifiedBy>Editor</c:lastModifiedBy>');

    const beforeAbsentClear = packageSnapshot(pkg);
    replaceModifiedAt(pkg, undefined);
    expect(packageSnapshot(pkg)).toEqual(beforeAbsentClear);
  });

  it('repairs type state and safely expands self-closing state', () => {
    const numericEntity = corePackage(coreXml(
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-30T01&#58;02&#58;03Z</dcterms:modified>',
    ));
    const beforeEntityNoOp = packageSnapshot(numericEntity);
    replaceModifiedAt(numericEntity, '2026-07-30T01:02:03Z');
    expect(packageSnapshot(numericEntity)).toEqual(beforeEntityNoOp);

    const missingType = corePackage(coreXml(
      '<d:modified custom="KEEP">2026-07-30T01:02:03Z</d:modified>',
      'c:coreProperties',
      `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DCTERMS_NAMESPACE}" xmlns:i="${XSI_NAMESPACE}"`,
    ));
    replaceModifiedAt(missingType, '2026-07-30T01:02:03Z');
    expect(partText(missingType)).toContain(
      '<d:modified custom="KEEP" i:type="d:W3CDTF">2026-07-30T01:02:03Z</d:modified>',
    );

    const wrongType = corePackage(coreXml(
      '<d:modified i:type="d:Other">invalid</d:modified>',
      'c:coreProperties',
      `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DCTERMS_NAMESPACE}" xmlns:i="${XSI_NAMESPACE}"`,
    ));
    replaceModifiedAt(wrongType, '2024-02-29T12:34:56Z');
    expect(partText(wrongType)).toContain(
      '<d:modified i:type="d:W3CDTF">2024-02-29T12:34:56Z</d:modified>',
    );

    const wrongNamespaceType = corePackage(coreXml(
      '<d:modified x:type="KEEP" xmlns:x="urn:wrong">2026-07-30T01:02:03Z</d:modified>',
      'c:coreProperties',
      `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DCTERMS_NAMESPACE}"`,
    ));
    replaceModifiedAt(wrongNamespaceType, '2026-07-30T01:02:03Z');
    expect(partText(wrongNamespaceType)).toContain('x:type="KEEP"');
    expect(partText(wrongNamespaceType)).toContain(`xmlns:xsi="${XSI_NAMESPACE}"`);
    expect(partText(wrongNamespaceType)).toContain('xsi:type="d:W3CDTF"');

    const propertyLocal = corePackage(coreXml(
      `<d:modified xmlns:d="${DCTERMS_NAMESPACE}" xmlns:i="${XSI_NAMESPACE}" i:type="d:Other">2026-07-30T01:02:03Z</d:modified>`,
      'c:coreProperties',
      `xmlns:c="${CORE_NAMESPACE}"`,
    ));
    replaceModifiedAt(propertyLocal, '2026-07-30T01:02:03Z');
    expect(readPresentationModifiedAt(propertyLocal)).toBe('2026-07-30T01:02:03Z');
    expect(partText(propertyLocal)).toContain('i:type="d:W3CDTF"');

    const selfClosingRoot = corePackage(
      `<?xml version="1.0"?><c:coreProperties xmlns:c="${CORE_NAMESPACE}"/>`,
    );
    replaceModifiedAt(selfClosingRoot, '2026-07-30T01:02:03Z');
    expect(partText(selfClosingRoot)).toContain(
      `<dcterms:modified xmlns:dcterms="${DCTERMS_NAMESPACE}" xmlns:xsi="${XSI_NAMESPACE}" xsi:type="dcterms:W3CDTF">2026-07-30T01:02:03Z</dcterms:modified>`,
    );

    const selfClosingProperty = corePackage(coreXml(
      '<d:modified custom="KEEP"/>',
      'c:coreProperties',
      `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DCTERMS_NAMESPACE}" xmlns:i="${XSI_NAMESPACE}"`,
    ));
    replaceModifiedAt(selfClosingProperty, '2026-07-30T01:02:03Z');
    expect(partText(selfClosingProperty)).toContain(
      '<d:modified custom="KEEP" i:type="d:W3CDTF">2026-07-30T01:02:03Z</d:modified>',
    );
  });

  it('rejects unsafe ownership without mutation', () => {
    const unsafe: OpcPackage[] = [
      corePackage(coreXml(
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-30T01:02:03Z</dcterms:modified>'
          + '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-31T01:02:03Z</dcterms:modified>',
      )),
      corePackage(coreXml(
        `<dcterms:modified xmlns:i="${XSI_NAMESPACE}" i:type="dcterms:W3CDTF" xsi:type="dcterms:W3CDTF">2026-07-30T01:02:03Z</dcterms:modified>`,
      )),
      corePackage(coreXml(
        '<dcterms:modified xsi:type="dcterms:W3CDTF">Before<x:keep xmlns:x="urn:test"/>After</dcterms:modified>',
      )),
      corePackage(coreXml(
        '<dcterms:modified xsi:type="dcterms:W3CDTF"><![CDATA[2026-07-30T01:02:03Z]]></dcterms:modified>',
      )),
      corePackage('<x:notCore xmlns:x="urn:test"/>'),
      corePackage(coreXml(
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-30T01:02:03Z</dcterms:modified>',
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
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-30T01:02:03Z</dcterms:modified>',
    ));
    duplicate.transaction(() => {
      duplicate.setPart(
        '/docProps/core2.xml',
        coreXml('<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-31T01:02:03Z</dcterms:modified>'),
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
      expect(() => replaceModifiedAt(pkg, '2026-07-30T01:02:03Z')).toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
      expect(() => replaceModifiedAt(pkg, undefined)).toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
    }
  });

  it('rolls back minimal creation and existing modified-at edits', () => {
    const created = OpcPackage.create();
    const createdBefore = packageSnapshot(created);
    expect(() => created.transaction(() => {
      replacePresentationModifiedAt(created, '2026-07-30T01:02:03Z');
      expect(readPresentationModifiedAt(created)).toBe('2026-07-30T01:02:03Z');
      throw new Error('restore modified-at creation');
    })).toThrow('restore modified-at creation');
    expect(packageSnapshot(created)).toEqual(createdBefore);

    const existing = corePackage(coreXml(
      '<dcterms:created xsi:type="dcterms:W3CDTF">2024-02-29T12:34:56Z</dcterms:created>'
        + '<dcterms:modified xsi:type="dcterms:W3CDTF">2024-02-29T13:34:56Z</dcterms:modified>',
    ));
    const existingBefore = packageSnapshot(existing);
    expect(() => existing.transaction(() => {
      replacePresentationModifiedAt(existing, '2026-07-30T01:02:03Z');
      expect(readPresentationModifiedAt(existing)).toBe('2026-07-30T01:02:03Z');
      throw new Error('restore modified-at edit');
    })).toThrow('restore modified-at edit');
    expect(packageSnapshot(existing)).toEqual(existingBefore);
    expect(readPresentationModifiedAt(existing)).toBe('2024-02-29T13:34:56Z');
    expect(readPresentationCreatedAt(existing)).toBe('2024-02-29T12:34:56Z');
  });
});
