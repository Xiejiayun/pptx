import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import {
  readPresentationLastModifiedBy,
  replacePresentationLastModifiedBy,
} from './presentation-last-modified-by.internal.js';

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

function replaceLastModifiedBy(pkg: OpcPackage, value: unknown): void {
  pkg.transaction(() => {
    replacePresentationLastModifiedBy(pkg, value as never);
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

describe('presentation last modified by core property', () => {
  it('reads only one namespace-correct direct simple lastModifiedBy without mutation', () => {
    const cases: readonly [OpcPackage, string | undefined][] = [
      [corePackage(), undefined],
      [corePackage(coreXml('<dc:creator>Creator only</dc:creator>')), undefined],
      [corePackage(coreXml(
        '<c:lastModifiedBy>编辑者 &amp; Reviewer</c:lastModifiedBy><c:revision>7</c:revision>',
        'c:coreProperties',
        `xmlns:c="${CORE_NAMESPACE}"`,
      ), { uri: '/metadata/core7.xml' }), '编辑者 & Reviewer'],
      [corePackage(coreXml('<cp:lastModifiedBy/>')), ''],
      [corePackage(coreXml(
        '<x:lastModifiedBy xmlns:x="urn:wrong">Wrong</x:lastModifiedBy>',
      )), undefined],
      [corePackage(coreXml(
        '<x:opaque xmlns:x="urn:test"><cp:lastModifiedBy>Nested</cp:lastModifiedBy></x:opaque>',
      )), undefined],
      [corePackage(coreXml(
        '<cp:lastModifiedBy>One</cp:lastModifiedBy><cp:lastModifiedBy>Two</cp:lastModifiedBy>',
      )), undefined],
      [corePackage(coreXml(
        '<cp:lastModifiedBy>Before<x:keep xmlns:x="urn:test"/>After</cp:lastModifiedBy>',
      )), undefined],
      [corePackage(coreXml(
        '<cp:lastModifiedBy><![CDATA[Raw]]></cp:lastModifiedBy>',
      )), undefined],
      [corePackage(
        `<x:notCore xmlns:x="urn:test"><cp:lastModifiedBy xmlns:cp="${CORE_NAMESPACE}">Wrong root</cp:lastModifiedBy></x:notCore>`,
      ), undefined],
      [corePackage(coreXml('<cp:lastModifiedBy>Wrong type</cp:lastModifiedBy>'), {
        contentType: 'application/xml',
      }), undefined],
    ];

    for (const [pkg, expected] of cases) {
      const before = packageSnapshot(pkg);
      expect(readPresentationLastModifiedBy(pkg)).toBe(expected);
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    const dangling = OpcPackage.create();
    setRootRelationships(
      dangling,
      `<Relationship Id="rId1" Type="${CORE_RELATIONSHIP}" Target="docProps/missing.xml"/>`,
    );
    const danglingBefore = packageSnapshot(dangling);
    expect(readPresentationLastModifiedBy(dangling)).toBeUndefined();
    expect(packageSnapshot(dangling)).toEqual(danglingBefore);

    const external = OpcPackage.create();
    setRootRelationships(
      external,
      `<Relationship Id="rId1" Type="${CORE_RELATIONSHIP}" Target="https://example.com/core.xml" TargetMode="External"/>`,
    );
    const externalBefore = packageSnapshot(external);
    expect(readPresentationLastModifiedBy(external)).toBeUndefined();
    expect(packageSnapshot(external)).toEqual(externalBefore);

    const duplicate = corePackage(coreXml(
      '<cp:lastModifiedBy>One</cp:lastModifiedBy>',
    ));
    duplicate.transaction(() => {
      duplicate.setPart(
        '/docProps/core2.xml',
        coreXml('<cp:lastModifiedBy>Two</cp:lastModifiedBy>'),
        CORE_CONTENT_TYPE,
      );
      duplicate.addRelationship('/', {
        id: 'rId2',
        type: CORE_RELATIONSHIP,
        target: 'docProps/core2.xml',
      });
    });
    const duplicateBefore = packageSnapshot(duplicate);
    expect(readPresentationLastModifiedBy(duplicate)).toBeUndefined();
    expect(packageSnapshot(duplicate)).toEqual(duplicateBefore);
  });

  it('creates only a minimal lastModifiedBy part without duplicate namespace declarations', () => {
    const pkg = OpcPackage.create();
    replaceLastModifiedBy(pkg, '作者 Alice & <Bob>');

    const relationship = pkg.relationships('/').find(({ type }) => type === CORE_RELATIONSHIP);
    expect(relationship).toMatchObject({
      target: 'docProps/core.xml',
      targetMode: 'Internal',
      resolvedTarget: '/docProps/core.xml',
    });
    expect(pkg.requirePart('/docProps/core.xml').contentType).toBe(CORE_CONTENT_TYPE);
    expect(partText(pkg)).toBe(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="${CORE_NAMESPACE}"><cp:lastModifiedBy>作者 Alice &amp; &lt;Bob&gt;</cp:lastModifiedBy></cp:coreProperties>`,
    );
    expect(partText(pkg).match(/xmlns:cp=/g)).toHaveLength(1);
    expect(readPresentationLastModifiedBy(pkg)).toBe('作者 Alice & <Bob>');

    const occupied = OpcPackage.create();
    const orphan = '<x:orphan xmlns:x="urn:test">KEEP</x:orphan>';
    occupied.setPart('/docProps/core.xml', orphan, 'application/xml');
    replaceLastModifiedBy(occupied, 'Allocated');
    const allocated = occupied.relationships('/').find(({ type }) => type === CORE_RELATIONSHIP);
    expect(allocated?.resolvedTarget).toBe('/docProps/core1.xml');
    expect(partText(occupied)).toBe(orphan);
    expect(readPresentationLastModifiedBy(occupied)).toBe('Allocated');
    expect(partText(occupied, '/docProps/core1.xml')).not.toContain('creator');
  });

  it('losslessly inserts, replaces, empties, and clears only lastModifiedBy', () => {
    const pkg = corePackage(coreXml(
      '\n  <dc:title>Quarterly</dc:title><dc:subject>Forecast</dc:subject>\n'
        + '  <dc:creator>Alice</dc:creator><cp:revision>7</cp:revision>\n'
        + '  <dcterms:created xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>\n'
        + '  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-30T01:00:00Z</dcterms:modified>\n'
        + '  <!--KEEP--><x:opaque xmlns:x="urn:test">KEEP</x:opaque>\n',
      'c:coreProperties',
      `xmlns:c="${CORE_NAMESPACE}" xmlns:dc="${DUBLIN_CORE_NAMESPACE}" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"`,
    ));
    pkg.setPart('/custom/keep.bin', new Uint8Array([1, 2, 3]), 'application/octet-stream');
    const unrelatedBefore = [...pkg.requirePart('/custom/keep.bin').bytes];

    replaceLastModifiedBy(pkg, 'Inserted & <safe>');
    let xml = partText(pkg);
    expect(xml).toContain('<c:lastModifiedBy>Inserted &amp; &lt;safe&gt;</c:lastModifiedBy>');
    expect(xml).toContain('<dc:title>Quarterly</dc:title>');
    expect(xml).toContain('<dc:subject>Forecast</dc:subject>');
    expect(xml).toContain('<dc:creator>Alice</dc:creator>');
    expect(xml).toContain('<cp:revision>7</cp:revision>');
    expect(xml).toContain('2026-07-30T00:00:00Z');
    expect(xml).toContain('2026-07-30T01:00:00Z');
    expect(xml).toContain('<!--KEEP-->');
    expect(xml).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    expect([...pkg.requirePart('/custom/keep.bin').bytes]).toEqual(unrelatedBefore);

    const beforeSame = packageSnapshot(pkg);
    replaceLastModifiedBy(pkg, 'Inserted & <safe>');
    expect(packageSnapshot(pkg)).toEqual(beforeSame);

    replaceLastModifiedBy(pkg, 'Replacement');
    expect(readPresentationLastModifiedBy(pkg)).toBe('Replacement');
    expect(partText(pkg)).toContain('<c:lastModifiedBy>Replacement</c:lastModifiedBy>');
    expect(partText(pkg)).toContain('<dc:creator>Alice</dc:creator>');

    replaceLastModifiedBy(pkg, '');
    expect(readPresentationLastModifiedBy(pkg)).toBe('');
    expect(partText(pkg)).toContain('<c:lastModifiedBy></c:lastModifiedBy>');

    replaceLastModifiedBy(pkg, undefined);
    expect(readPresentationLastModifiedBy(pkg)).toBeUndefined();
    xml = partText(pkg);
    expect(xml).not.toContain('<c:lastModifiedBy');
    expect(xml).toContain('<dc:creator>Alice</dc:creator>');
    expect(xml).toContain('<cp:revision>7</cp:revision>');
    expect(pkg.relationships('/').filter(({ type }) => type === CORE_RELATIONSHIP)).toHaveLength(1);

    const beforeAbsentClear = packageSnapshot(pkg);
    replaceLastModifiedBy(pkg, undefined);
    expect(packageSnapshot(pkg)).toEqual(beforeAbsentClear);
  });

  it('preserves semantic no-ops and safely expands self-closing state', () => {
    const numericEntity = corePackage(coreXml(
      '<cp:lastModifiedBy>A&#38;B</cp:lastModifiedBy><dc:creator>Alice</dc:creator>',
    ));
    const beforeEntityNoOp = packageSnapshot(numericEntity);
    replaceLastModifiedBy(numericEntity, 'A&B');
    expect(packageSnapshot(numericEntity)).toEqual(beforeEntityNoOp);

    const selfClosingRoot = corePackage(
      `<?xml version="1.0"?><c:coreProperties xmlns:c="${CORE_NAMESPACE}"/>`,
    );
    replaceLastModifiedBy(selfClosingRoot, 'Root expanded');
    expect(readPresentationLastModifiedBy(selfClosingRoot)).toBe('Root expanded');
    expect(partText(selfClosingRoot)).toContain(
      '<c:lastModifiedBy>Root expanded</c:lastModifiedBy>',
    );

    const selfClosingProperty = corePackage(coreXml(
      '<cp:lastModifiedBy custom="KEEP"/><dc:creator>Alice</dc:creator>',
    ));
    replaceLastModifiedBy(selfClosingProperty, 'Expanded');
    expect(partText(selfClosingProperty)).toContain(
      '<cp:lastModifiedBy custom="KEEP">Expanded</cp:lastModifiedBy>',
    );
    expect(partText(selfClosingProperty)).toContain('<dc:creator>Alice</dc:creator>');

    const wrongNamespace = corePackage(coreXml(
      '<x:lastModifiedBy xmlns:x="urn:wrong">KEEP</x:lastModifiedBy>',
      'cp:coreProperties',
      `xmlns:cp="${CORE_NAMESPACE}"`,
    ));
    replaceLastModifiedBy(wrongNamespace, 'Direct');
    expect(partText(wrongNamespace)).toContain(
      '<x:lastModifiedBy xmlns:x="urn:wrong">KEEP</x:lastModifiedBy>',
    );
    expect(partText(wrongNamespace)).toContain(
      '<cp:lastModifiedBy>Direct</cp:lastModifiedBy>',
    );
  });

  it('rejects invalid values and unsafe ownership without mutation', () => {
    const invalidValues = [
      null,
      true,
      false,
      0,
      1n,
      {},
      [],
      Symbol('lastModifiedBy'),
      'bad\u0001editor',
    ] as const;
    for (const value of invalidValues) {
      const pkg = OpcPackage.create();
      const before = packageSnapshot(pkg);
      expect(() => replaceLastModifiedBy(pkg, value)).toThrow(TypeError);
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    const unsafe: OpcPackage[] = [
      corePackage(coreXml(
        '<cp:lastModifiedBy>One</cp:lastModifiedBy><cp:lastModifiedBy>Two</cp:lastModifiedBy>',
      )),
      corePackage(coreXml(
        '<cp:lastModifiedBy>Before<x:keep xmlns:x="urn:test"/>After</cp:lastModifiedBy>',
      )),
      corePackage(coreXml('<cp:lastModifiedBy><![CDATA[Raw]]></cp:lastModifiedBy>')),
      corePackage('<x:notCore xmlns:x="urn:test"/>'),
      corePackage(coreXml('<cp:lastModifiedBy>Wrong type</cp:lastModifiedBy>'), {
        contentType: 'application/xml',
      }),
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
      '<cp:lastModifiedBy>One</cp:lastModifiedBy>',
    ));
    duplicate.transaction(() => {
      duplicate.setPart(
        '/docProps/core2.xml',
        coreXml('<cp:lastModifiedBy>Two</cp:lastModifiedBy>'),
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
      expect(() => replaceLastModifiedBy(pkg, 'Replacement')).toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
    }
  });

  it('rolls back minimal creation and existing lastModifiedBy edits', () => {
    const created = OpcPackage.create();
    const createdBefore = packageSnapshot(created);

    expect(() => created.transaction(() => {
      replacePresentationLastModifiedBy(created, 'Temporary');
      expect(readPresentationLastModifiedBy(created)).toBe('Temporary');
      throw new Error('restore lastModifiedBy creation');
    })).toThrow('restore lastModifiedBy creation');
    expect(packageSnapshot(created)).toEqual(createdBefore);
    expect(created.hasPart('/docProps/core.xml')).toBe(false);
    expect(created.relationships('/')).toHaveLength(0);

    const existing = corePackage(coreXml(
      '<cp:lastModifiedBy>Original</cp:lastModifiedBy><dc:creator>Alice</dc:creator>',
    ));
    const existingBefore = packageSnapshot(existing);
    expect(() => existing.transaction(() => {
      replacePresentationLastModifiedBy(existing, 'Temporary');
      expect(readPresentationLastModifiedBy(existing)).toBe('Temporary');
      throw new Error('restore lastModifiedBy edit');
    })).toThrow('restore lastModifiedBy edit');
    expect(packageSnapshot(existing)).toEqual(existingBefore);
    expect(readPresentationLastModifiedBy(existing)).toBe('Original');
    expect(partText(existing)).toContain('<dc:creator>Alice</dc:creator>');
  });
});
