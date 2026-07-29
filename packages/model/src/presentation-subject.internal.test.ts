import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import {
  readPresentationSubject,
  replacePresentationSubject,
} from './presentation-subject.internal.js';

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

function replaceSubject(pkg: OpcPackage, value: unknown): void {
  pkg.transaction(() => {
    replacePresentationSubject(pkg, value as never);
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

describe('presentation subject core property', () => {
  it('reads only one namespace-correct direct simple subject without mutation', () => {
    const cases: readonly [OpcPackage, string | undefined][] = [
      [corePackage(), undefined],
      [corePackage(coreXml('<dc:title>Title only</dc:title>')), undefined],
      [corePackage(coreXml(
        '<d:subject>主题 &amp; Forecast</d:subject><cp:revision>7</cp:revision>',
        'c:coreProperties',
        `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DUBLIN_CORE_NAMESPACE}" xmlns:cp="${CORE_NAMESPACE}"`,
      ), { uri: '/metadata/core7.xml' }), '主题 & Forecast'],
      [corePackage(coreXml('<dc:subject/>')), ''],
      [corePackage(coreXml('<x:subject xmlns:x="urn:wrong">Wrong</x:subject>')), undefined],
      [corePackage(coreXml(
        '<x:opaque xmlns:x="urn:test"><dc:subject>Nested</dc:subject></x:opaque>',
      )), undefined],
      [corePackage(coreXml(
        '<dc:subject>One</dc:subject><dc:subject>Two</dc:subject>',
      )), undefined],
      [corePackage(coreXml(
        '<dc:subject>Before<x:keep xmlns:x="urn:test"/>After</dc:subject>',
      )), undefined],
      [corePackage(coreXml('<dc:subject><![CDATA[Raw]]></dc:subject>')), undefined],
      [corePackage(
        `<x:notCore xmlns:x="urn:test"><dc:subject xmlns:dc="${DUBLIN_CORE_NAMESPACE}">Wrong root</dc:subject></x:notCore>`,
      ), undefined],
      [corePackage(coreXml('<dc:subject>Wrong type</dc:subject>'), {
        contentType: 'application/xml',
      }), undefined],
    ];

    for (const [pkg, expected] of cases) {
      const before = packageSnapshot(pkg);
      expect(readPresentationSubject(pkg)).toBe(expected);
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    const dangling = OpcPackage.create();
    setRootRelationships(
      dangling,
      `<Relationship Id="rId1" Type="${CORE_RELATIONSHIP}" Target="docProps/missing.xml"/>`,
    );
    const danglingBefore = packageSnapshot(dangling);
    expect(readPresentationSubject(dangling)).toBeUndefined();
    expect(packageSnapshot(dangling)).toEqual(danglingBefore);

    const external = OpcPackage.create();
    setRootRelationships(
      external,
      `<Relationship Id="rId1" Type="${CORE_RELATIONSHIP}" Target="https://example.com/core.xml" TargetMode="External"/>`,
    );
    const externalBefore = packageSnapshot(external);
    expect(readPresentationSubject(external)).toBeUndefined();
    expect(packageSnapshot(external)).toEqual(externalBefore);

    const duplicate = corePackage(coreXml('<dc:subject>One</dc:subject>'));
    duplicate.transaction(() => {
      duplicate.setPart(
        '/docProps/core2.xml',
        coreXml('<dc:subject>Two</dc:subject>'),
        CORE_CONTENT_TYPE,
      );
      duplicate.addRelationship('/', {
        id: 'rId2',
        type: CORE_RELATIONSHIP,
        target: 'docProps/core2.xml',
      });
    });
    const duplicateBefore = packageSnapshot(duplicate);
    expect(readPresentationSubject(duplicate)).toBeUndefined();
    expect(packageSnapshot(duplicate)).toEqual(duplicateBefore);
  });

  it('creates only a minimal subject part and avoids an occupied canonical URI', () => {
    const pkg = OpcPackage.create();
    replaceSubject(pkg, '主题 & <Forecast>');

    const relationship = pkg.relationships('/').find(({ type }) => type === CORE_RELATIONSHIP);
    expect(relationship).toMatchObject({
      target: 'docProps/core.xml',
      targetMode: 'Internal',
      resolvedTarget: '/docProps/core.xml',
    });
    expect(pkg.requirePart('/docProps/core.xml').contentType).toBe(CORE_CONTENT_TYPE);
    expect(partText(pkg)).toBe(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="${CORE_NAMESPACE}" xmlns:dc="${DUBLIN_CORE_NAMESPACE}"><dc:subject>主题 &amp; &lt;Forecast&gt;</dc:subject></cp:coreProperties>`,
    );
    expect(readPresentationSubject(pkg)).toBe('主题 & <Forecast>');

    const occupied = OpcPackage.create();
    const orphan = '<x:orphan xmlns:x="urn:test">KEEP</x:orphan>';
    occupied.setPart('/docProps/core.xml', orphan, 'application/xml');
    replaceSubject(occupied, 'Allocated');
    const allocated = occupied.relationships('/').find(({ type }) => type === CORE_RELATIONSHIP);
    expect(allocated?.resolvedTarget).toBe('/docProps/core1.xml');
    expect(partText(occupied)).toBe(orphan);
    expect(readPresentationSubject(occupied)).toBe('Allocated');
    expect(partText(occupied, '/docProps/core1.xml')).not.toContain('creator');
  });

  it('losslessly inserts, replaces, empties, and clears only the subject', () => {
    const pkg = corePackage(coreXml(
      '\n  <d:title>Quarterly</d:title><d:creator>Alice</d:creator>\n'
        + '  <cp:lastModifiedBy>Editor</cp:lastModifiedBy><cp:revision>7</cp:revision>\n'
        + '  <dcterms:created xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>\n'
        + '  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-30T01:00:00Z</dcterms:modified>\n'
        + '  <!--KEEP--><x:opaque xmlns:x="urn:test">KEEP</x:opaque>\n',
      'c:coreProperties',
      `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DUBLIN_CORE_NAMESPACE}" xmlns:cp="${CORE_NAMESPACE}" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"`,
    ));
    pkg.setPart('/custom/keep.bin', new Uint8Array([1, 2, 3]), 'application/octet-stream');
    const unrelatedBefore = [...pkg.requirePart('/custom/keep.bin').bytes];

    replaceSubject(pkg, 'Inserted & <safe>');
    let xml = partText(pkg);
    expect(xml).toContain('<d:subject>Inserted &amp; &lt;safe&gt;</d:subject>');
    expect(xml).toContain('<d:title>Quarterly</d:title><d:creator>Alice</d:creator>');
    expect(xml).toContain('<cp:lastModifiedBy>Editor</cp:lastModifiedBy>');
    expect(xml).toContain('<cp:revision>7</cp:revision>');
    expect(xml).toContain(
      '<dcterms:created xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>',
    );
    expect(xml).toContain(
      '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-30T01:00:00Z</dcterms:modified>',
    );
    expect(xml).toContain('<!--KEEP-->');
    expect(xml).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    expect([...pkg.requirePart('/custom/keep.bin').bytes]).toEqual(unrelatedBefore);

    const beforeSame = packageSnapshot(pkg);
    replaceSubject(pkg, 'Inserted & <safe>');
    expect(packageSnapshot(pkg)).toEqual(beforeSame);

    replaceSubject(pkg, 'Replacement');
    expect(readPresentationSubject(pkg)).toBe('Replacement');
    expect(partText(pkg)).toContain('<d:subject>Replacement</d:subject>');
    expect(partText(pkg)).toContain('<d:creator>Alice</d:creator>');

    replaceSubject(pkg, '');
    expect(readPresentationSubject(pkg)).toBe('');
    expect(partText(pkg)).toContain('<d:subject></d:subject>');

    replaceSubject(pkg, undefined);
    expect(readPresentationSubject(pkg)).toBeUndefined();
    xml = partText(pkg);
    expect(xml).not.toContain('<d:subject');
    expect(xml).toContain('<d:title>Quarterly</d:title><d:creator>Alice</d:creator>');
    expect(xml).toContain('<cp:lastModifiedBy>Editor</cp:lastModifiedBy>');
    expect(pkg.relationships('/').filter(({ type }) => type === CORE_RELATIONSHIP)).toHaveLength(1);

    const beforeAbsentClear = packageSnapshot(pkg);
    replaceSubject(pkg, undefined);
    expect(packageSnapshot(pkg)).toEqual(beforeAbsentClear);
  });

  it('preserves semantic no-ops and safely expands self-closing state', () => {
    const numericEntity = corePackage(coreXml(
      '<dc:subject>A&#38;B</dc:subject><cp:revision>7</cp:revision>',
    ));
    const beforeEntityNoOp = packageSnapshot(numericEntity);
    replaceSubject(numericEntity, 'A&B');
    expect(packageSnapshot(numericEntity)).toEqual(beforeEntityNoOp);

    const selfClosingRoot = corePackage(
      `<?xml version="1.0"?><c:coreProperties xmlns:c="${CORE_NAMESPACE}"/>`,
    );
    replaceSubject(selfClosingRoot, 'Root expanded');
    expect(readPresentationSubject(selfClosingRoot)).toBe('Root expanded');
    expect(partText(selfClosingRoot)).toContain(
      `<dc:subject xmlns:dc="${DUBLIN_CORE_NAMESPACE}">Root expanded</dc:subject>`,
    );

    const selfClosingSubject = corePackage(coreXml(
      '<dc:subject custom="KEEP"/><cp:revision>7</cp:revision>',
    ));
    replaceSubject(selfClosingSubject, 'Expanded');
    expect(partText(selfClosingSubject)).toContain(
      '<dc:subject custom="KEEP">Expanded</dc:subject>',
    );
    expect(partText(selfClosingSubject)).toContain('<cp:revision>7</cp:revision>');

    const wrongNamespace = corePackage(coreXml(
      '<x:subject xmlns:x="urn:wrong">KEEP</x:subject>',
      'cp:coreProperties',
      `xmlns:cp="${CORE_NAMESPACE}"`,
    ));
    replaceSubject(wrongNamespace, 'Direct');
    expect(partText(wrongNamespace)).toContain(
      '<x:subject xmlns:x="urn:wrong">KEEP</x:subject>',
    );
    expect(partText(wrongNamespace)).toContain(
      `<dc:subject xmlns:dc="${DUBLIN_CORE_NAMESPACE}">Direct</dc:subject>`,
    );
  });

  it('rejects invalid values and unsafe ownership without mutation', () => {
    for (const value of [null, false, 0, {}, [], Symbol('subject'), 'bad\u0001subject']) {
      const pkg = OpcPackage.create();
      const before = packageSnapshot(pkg);
      expect(() => replaceSubject(pkg, value)).toThrow(TypeError);
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    const unsafe: OpcPackage[] = [
      corePackage(coreXml(
        '<dc:subject>One</dc:subject><dc:subject>Two</dc:subject>',
      )),
      corePackage(coreXml(
        '<dc:subject>Before<x:keep xmlns:x="urn:test"/>After</dc:subject>',
      )),
      corePackage(coreXml('<dc:subject><![CDATA[Raw]]></dc:subject>')),
      corePackage('<x:notCore xmlns:x="urn:test"/>'),
      corePackage(coreXml('<dc:subject>Wrong type</dc:subject>'), {
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

    const duplicate = corePackage(coreXml('<dc:subject>One</dc:subject>'));
    duplicate.transaction(() => {
      duplicate.setPart(
        '/docProps/core2.xml',
        coreXml('<dc:subject>Two</dc:subject>'),
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
      expect(() => replaceSubject(pkg, 'Replacement')).toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
    }
  });

  it('rolls back minimal creation and existing subject edits', () => {
    const created = OpcPackage.create();
    const createdBefore = packageSnapshot(created);

    expect(() => created.transaction(() => {
      replacePresentationSubject(created, 'Temporary');
      expect(readPresentationSubject(created)).toBe('Temporary');
      throw new Error('restore subject creation');
    })).toThrow('restore subject creation');

    expect(packageSnapshot(created)).toEqual(createdBefore);
    expect(created.hasPart('/docProps/core.xml')).toBe(false);
    expect(created.relationships('/')).toHaveLength(0);

    const existing = corePackage(coreXml(
      '<dc:subject>Original</dc:subject><dc:creator>Alice</dc:creator>',
    ));
    const existingBefore = packageSnapshot(existing);
    expect(() => existing.transaction(() => {
      replacePresentationSubject(existing, 'Temporary');
      expect(readPresentationSubject(existing)).toBe('Temporary');
      throw new Error('restore subject edit');
    })).toThrow('restore subject edit');
    expect(packageSnapshot(existing)).toEqual(existingBefore);
    expect(readPresentationSubject(existing)).toBe('Original');
    expect(partText(existing)).toContain('<dc:creator>Alice</dc:creator>');
  });
});
