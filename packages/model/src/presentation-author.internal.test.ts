import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import {
  readPresentationAuthor,
  replacePresentationAuthor,
} from './presentation-author.internal.js';

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

function replaceAuthor(pkg: OpcPackage, value: unknown): void {
  pkg.transaction(() => {
    replacePresentationAuthor(pkg, value as never);
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

describe('presentation author core property', () => {
  it('reads only one namespace-correct direct simple creator without mutation', () => {
    const cases: readonly [OpcPackage, string | undefined][] = [
      [corePackage(), undefined],
      [corePackage(coreXml('<dc:title>Title only</dc:title>')), undefined],
      [corePackage(coreXml(
        '<d:creator>作者 Alice &amp; Bob</d:creator><cp:lastModifiedBy>Editor</cp:lastModifiedBy>',
        'c:coreProperties',
        `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DUBLIN_CORE_NAMESPACE}" xmlns:cp="${CORE_NAMESPACE}"`,
      ), { uri: '/metadata/core7.xml' }), '作者 Alice & Bob'],
      [corePackage(coreXml('<dc:creator/>')), ''],
      [corePackage(coreXml('<x:creator xmlns:x="urn:wrong">Wrong</x:creator>')), undefined],
      [corePackage(coreXml(
        '<x:opaque xmlns:x="urn:test"><dc:creator>Nested</dc:creator></x:opaque>',
      )), undefined],
      [corePackage(coreXml(
        '<dc:creator>One</dc:creator><dc:creator>Two</dc:creator>',
      )), undefined],
      [corePackage(coreXml(
        '<dc:creator>Before<x:keep xmlns:x="urn:test"/>After</dc:creator>',
      )), undefined],
      [corePackage(coreXml('<dc:creator><![CDATA[Raw]]></dc:creator>')), undefined],
      [corePackage(
        `<x:notCore xmlns:x="urn:test"><dc:creator xmlns:dc="${DUBLIN_CORE_NAMESPACE}">Wrong root</dc:creator></x:notCore>`,
      ), undefined],
      [corePackage(coreXml('<dc:creator>Wrong type</dc:creator>'), {
        contentType: 'application/xml',
      }), undefined],
    ];

    for (const [pkg, expected] of cases) {
      const before = packageSnapshot(pkg);
      expect(readPresentationAuthor(pkg)).toBe(expected);
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    const dangling = OpcPackage.create();
    setRootRelationships(
      dangling,
      `<Relationship Id="rId1" Type="${CORE_RELATIONSHIP}" Target="docProps/missing.xml"/>`,
    );
    const danglingBefore = packageSnapshot(dangling);
    expect(readPresentationAuthor(dangling)).toBeUndefined();
    expect(packageSnapshot(dangling)).toEqual(danglingBefore);

    const external = OpcPackage.create();
    setRootRelationships(
      external,
      `<Relationship Id="rId1" Type="${CORE_RELATIONSHIP}" Target="https://example.com/core.xml" TargetMode="External"/>`,
    );
    const externalBefore = packageSnapshot(external);
    expect(readPresentationAuthor(external)).toBeUndefined();
    expect(packageSnapshot(external)).toEqual(externalBefore);

    const duplicate = corePackage(coreXml('<dc:creator>One</dc:creator>'));
    duplicate.transaction(() => {
      duplicate.setPart(
        '/docProps/core2.xml',
        coreXml('<dc:creator>Two</dc:creator>'),
        CORE_CONTENT_TYPE,
      );
      duplicate.addRelationship('/', {
        id: 'rId2',
        type: CORE_RELATIONSHIP,
        target: 'docProps/core2.xml',
      });
    });
    const duplicateBefore = packageSnapshot(duplicate);
    expect(readPresentationAuthor(duplicate)).toBeUndefined();
    expect(packageSnapshot(duplicate)).toEqual(duplicateBefore);
  });

  it('creates only a minimal creator part and avoids an occupied canonical URI', () => {
    const pkg = OpcPackage.create();
    replaceAuthor(pkg, '作者 Alice & <Bob>');

    const relationship = pkg.relationships('/').find(({ type }) => type === CORE_RELATIONSHIP);
    expect(relationship).toMatchObject({
      target: 'docProps/core.xml',
      targetMode: 'Internal',
      resolvedTarget: '/docProps/core.xml',
    });
    expect(pkg.requirePart('/docProps/core.xml').contentType).toBe(CORE_CONTENT_TYPE);
    expect(partText(pkg)).toBe(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="${CORE_NAMESPACE}" xmlns:dc="${DUBLIN_CORE_NAMESPACE}"><dc:creator>作者 Alice &amp; &lt;Bob&gt;</dc:creator></cp:coreProperties>`,
    );
    expect(readPresentationAuthor(pkg)).toBe('作者 Alice & <Bob>');

    const occupied = OpcPackage.create();
    const orphan = '<x:orphan xmlns:x="urn:test">KEEP</x:orphan>';
    occupied.setPart('/docProps/core.xml', orphan, 'application/xml');
    replaceAuthor(occupied, 'Allocated');
    const allocated = occupied.relationships('/').find(({ type }) => type === CORE_RELATIONSHIP);
    expect(allocated?.resolvedTarget).toBe('/docProps/core1.xml');
    expect(partText(occupied)).toBe(orphan);
    expect(readPresentationAuthor(occupied)).toBe('Allocated');
    expect(partText(occupied, '/docProps/core1.xml')).not.toContain('lastModifiedBy');
  });

  it('losslessly inserts, replaces, empties, and clears only the creator', () => {
    const pkg = corePackage(coreXml(
      '\n  <dc:title>Quarterly</dc:title>\n  <cp:lastModifiedBy>Editor</cp:lastModifiedBy>\n'
        + '  <dc:subject>Subject</dc:subject><cp:revision>7</cp:revision>\n'
        + '  <dcterms:created xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>\n'
        + '  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-30T01:00:00Z</dcterms:modified>\n'
        + '  <!--KEEP--><x:opaque xmlns:x="urn:test">KEEP</x:opaque>\n',
      'c:coreProperties',
      `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DUBLIN_CORE_NAMESPACE}" xmlns:dc="${DUBLIN_CORE_NAMESPACE}" xmlns:cp="${CORE_NAMESPACE}" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"`,
    ));
    pkg.setPart('/custom/keep.bin', new Uint8Array([1, 2, 3]), 'application/octet-stream');
    const unrelatedBefore = [...pkg.requirePart('/custom/keep.bin').bytes];

    replaceAuthor(pkg, 'Inserted & <safe>');
    let xml = partText(pkg);
    expect(xml).toContain('<d:creator>Inserted &amp; &lt;safe&gt;</d:creator>');
    expect(xml).toContain('<dc:title>Quarterly</dc:title>');
    expect(xml).toContain('<cp:lastModifiedBy>Editor</cp:lastModifiedBy>');
    expect(xml).toContain('<dc:subject>Subject</dc:subject><cp:revision>7</cp:revision>');
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
    replaceAuthor(pkg, 'Inserted & <safe>');
    expect(packageSnapshot(pkg)).toEqual(beforeSame);

    replaceAuthor(pkg, 'Replacement');
    expect(readPresentationAuthor(pkg)).toBe('Replacement');
    expect(partText(pkg)).toContain('<d:creator>Replacement</d:creator>');
    expect(partText(pkg)).toContain('<cp:lastModifiedBy>Editor</cp:lastModifiedBy>');

    replaceAuthor(pkg, '');
    expect(readPresentationAuthor(pkg)).toBe('');
    expect(partText(pkg)).toContain('<d:creator></d:creator>');
    expect(partText(pkg)).toContain('<cp:lastModifiedBy>Editor</cp:lastModifiedBy>');

    replaceAuthor(pkg, undefined);
    expect(readPresentationAuthor(pkg)).toBeUndefined();
    xml = partText(pkg);
    expect(xml).not.toContain('<d:creator');
    expect(xml).toContain('<dc:title>Quarterly</dc:title>');
    expect(xml).toContain('<cp:lastModifiedBy>Editor</cp:lastModifiedBy>');
    expect(pkg.relationships('/').filter(({ type }) => type === CORE_RELATIONSHIP)).toHaveLength(1);

    const beforeAbsentClear = packageSnapshot(pkg);
    replaceAuthor(pkg, undefined);
    expect(packageSnapshot(pkg)).toEqual(beforeAbsentClear);
  });

  it('preserves semantic no-ops and safely expands self-closing state', () => {
    const numericEntity = corePackage(coreXml(
      '<dc:creator>A&#38;B</dc:creator><cp:lastModifiedBy>Editor</cp:lastModifiedBy>',
    ));
    const beforeEntityNoOp = packageSnapshot(numericEntity);
    replaceAuthor(numericEntity, 'A&B');
    expect(packageSnapshot(numericEntity)).toEqual(beforeEntityNoOp);

    const selfClosingRoot = corePackage(
      `<?xml version="1.0"?><c:coreProperties xmlns:c="${CORE_NAMESPACE}"/>`,
    );
    replaceAuthor(selfClosingRoot, 'Root expanded');
    expect(readPresentationAuthor(selfClosingRoot)).toBe('Root expanded');
    expect(partText(selfClosingRoot)).toContain(
      `<dc:creator xmlns:dc="${DUBLIN_CORE_NAMESPACE}">Root expanded</dc:creator>`,
    );

    const selfClosingCreator = corePackage(coreXml(
      '<dc:creator custom="KEEP"/><cp:lastModifiedBy>Editor</cp:lastModifiedBy>',
    ));
    replaceAuthor(selfClosingCreator, 'Expanded');
    expect(partText(selfClosingCreator)).toContain(
      '<dc:creator custom="KEEP">Expanded</dc:creator>',
    );
    expect(partText(selfClosingCreator)).toContain(
      '<cp:lastModifiedBy>Editor</cp:lastModifiedBy>',
    );

    const wrongNamespace = corePackage(coreXml(
      '<x:creator xmlns:x="urn:wrong">KEEP</x:creator>',
      'cp:coreProperties',
      `xmlns:cp="${CORE_NAMESPACE}"`,
    ));
    replaceAuthor(wrongNamespace, 'Direct');
    expect(partText(wrongNamespace)).toContain(
      '<x:creator xmlns:x="urn:wrong">KEEP</x:creator>',
    );
    expect(partText(wrongNamespace)).toContain(
      `<dc:creator xmlns:dc="${DUBLIN_CORE_NAMESPACE}">Direct</dc:creator>`,
    );
  });

  it('rejects invalid values and unsafe ownership without mutation', () => {
    for (const value of [null, false, 0, {}, [], Symbol('author'), 'bad\u0001author']) {
      const pkg = OpcPackage.create();
      const before = packageSnapshot(pkg);
      expect(() => replaceAuthor(pkg, value)).toThrow(TypeError);
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    const unsafe: OpcPackage[] = [
      corePackage(coreXml(
        '<dc:creator>One</dc:creator><dc:creator>Two</dc:creator>',
      )),
      corePackage(coreXml(
        '<dc:creator>Before<x:keep xmlns:x="urn:test"/>After</dc:creator>',
      )),
      corePackage(coreXml('<dc:creator><![CDATA[Raw]]></dc:creator>')),
      corePackage('<x:notCore xmlns:x="urn:test"/>'),
      corePackage(coreXml('<dc:creator>Wrong type</dc:creator>'), {
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

    const duplicate = corePackage(coreXml('<dc:creator>One</dc:creator>'));
    duplicate.transaction(() => {
      duplicate.setPart(
        '/docProps/core2.xml',
        coreXml('<dc:creator>Two</dc:creator>'),
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
      expect(() => replaceAuthor(pkg, 'Replacement')).toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
    }
  });

  it('rolls back minimal creation and existing creator edits', () => {
    const created = OpcPackage.create();
    const createdBefore = packageSnapshot(created);

    expect(() => created.transaction(() => {
      replacePresentationAuthor(created, 'Temporary');
      expect(readPresentationAuthor(created)).toBe('Temporary');
      throw new Error('restore author creation');
    })).toThrow('restore author creation');

    expect(packageSnapshot(created)).toEqual(createdBefore);
    expect(created.hasPart('/docProps/core.xml')).toBe(false);
    expect(created.relationships('/')).toHaveLength(0);

    const existing = corePackage(coreXml(
      '<dc:creator>Original</dc:creator><cp:lastModifiedBy>Editor</cp:lastModifiedBy>',
    ));
    const existingBefore = packageSnapshot(existing);
    expect(() => existing.transaction(() => {
      replacePresentationAuthor(existing, 'Temporary');
      expect(readPresentationAuthor(existing)).toBe('Temporary');
      throw new Error('restore author edit');
    })).toThrow('restore author edit');
    expect(packageSnapshot(existing)).toEqual(existingBefore);
    expect(readPresentationAuthor(existing)).toBe('Original');
    expect(partText(existing)).toContain('<cp:lastModifiedBy>Editor</cp:lastModifiedBy>');
  });
});
