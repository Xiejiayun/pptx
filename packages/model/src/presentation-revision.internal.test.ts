import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import {
  readPresentationRevision,
  replacePresentationRevision,
} from './presentation-revision.internal.js';

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

function replaceRevision(pkg: OpcPackage, value: unknown): void {
  pkg.transaction(() => {
    replacePresentationRevision(pkg, value as never);
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

describe('presentation revision core property', () => {
  it('reads only one namespace-correct direct ASCII whole-number without mutation', () => {
    const cases: readonly [OpcPackage, string | undefined][] = [
      [corePackage(), undefined],
      [corePackage(coreXml('<dc:title>Title only</dc:title>')), undefined],
      [corePackage(coreXml('<cp:revision>0</cp:revision>')), '0'],
      [corePackage(coreXml('<cp:revision>007</cp:revision>')), '007'],
      [corePackage(coreXml('<cp:revision>123456789012345678901234567890</cp:revision>')),
        '123456789012345678901234567890'],
      [corePackage(coreXml(
        '<c:revision>42</c:revision><d:title>Quarterly</d:title>',
        'c:coreProperties',
        `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DUBLIN_CORE_NAMESPACE}"`,
      ), { uri: '/metadata/core7.xml' }), '42'],
      [corePackage(coreXml('<cp:revision/>')), undefined],
      [corePackage(coreXml('<cp:revision> </cp:revision>')), undefined],
      [corePackage(coreXml('<cp:revision>+1</cp:revision>')), undefined],
      [corePackage(coreXml('<cp:revision>-1</cp:revision>')), undefined],
      [corePackage(coreXml('<cp:revision>1.5</cp:revision>')), undefined],
      [corePackage(coreXml('<cp:revision>1e3</cp:revision>')), undefined],
      [corePackage(coreXml('<cp:revision>１２</cp:revision>')), undefined],
      [corePackage(coreXml('<x:revision xmlns:x="urn:wrong">7</x:revision>')), undefined],
      [corePackage(coreXml(
        '<x:opaque xmlns:x="urn:test"><cp:revision>7</cp:revision></x:opaque>',
      )), undefined],
      [corePackage(coreXml(
        '<cp:revision>1</cp:revision><cp:revision>2</cp:revision>',
      )), undefined],
      [corePackage(coreXml(
        '<cp:revision>1<x:keep xmlns:x="urn:test"/>2</cp:revision>',
      )), undefined],
      [corePackage(coreXml('<cp:revision><![CDATA[7]]></cp:revision>')), undefined],
      [corePackage(
        `<x:notCore xmlns:x="urn:test"><cp:revision xmlns:cp="${CORE_NAMESPACE}">7</cp:revision></x:notCore>`,
      ), undefined],
      [corePackage(coreXml('<cp:revision>7</cp:revision>'), {
        contentType: 'application/xml',
      }), undefined],
    ];

    for (const [pkg, expected] of cases) {
      const before = packageSnapshot(pkg);
      expect(readPresentationRevision(pkg)).toBe(expected);
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    const dangling = OpcPackage.create();
    setRootRelationships(
      dangling,
      `<Relationship Id="rId1" Type="${CORE_RELATIONSHIP}" Target="docProps/missing.xml"/>`,
    );
    const danglingBefore = packageSnapshot(dangling);
    expect(readPresentationRevision(dangling)).toBeUndefined();
    expect(packageSnapshot(dangling)).toEqual(danglingBefore);

    const external = OpcPackage.create();
    setRootRelationships(
      external,
      `<Relationship Id="rId1" Type="${CORE_RELATIONSHIP}" Target="https://example.com/core.xml" TargetMode="External"/>`,
    );
    const externalBefore = packageSnapshot(external);
    expect(readPresentationRevision(external)).toBeUndefined();
    expect(packageSnapshot(external)).toEqual(externalBefore);

    const duplicate = corePackage(coreXml('<cp:revision>1</cp:revision>'));
    duplicate.transaction(() => {
      duplicate.setPart(
        '/docProps/core2.xml',
        coreXml('<cp:revision>2</cp:revision>'),
        CORE_CONTENT_TYPE,
      );
      duplicate.addRelationship('/', {
        id: 'rId2',
        type: CORE_RELATIONSHIP,
        target: 'docProps/core2.xml',
      });
    });
    const duplicateBefore = packageSnapshot(duplicate);
    expect(readPresentationRevision(duplicate)).toBeUndefined();
    expect(packageSnapshot(duplicate)).toEqual(duplicateBefore);
  });

  it('creates only a minimal revision part without duplicate namespace declarations', () => {
    const pkg = OpcPackage.create();
    replaceRevision(pkg, '007');

    const relationship = pkg.relationships('/').find(({ type }) => type === CORE_RELATIONSHIP);
    expect(relationship).toMatchObject({
      target: 'docProps/core.xml',
      targetMode: 'Internal',
      resolvedTarget: '/docProps/core.xml',
    });
    expect(pkg.requirePart('/docProps/core.xml').contentType).toBe(CORE_CONTENT_TYPE);
    expect(partText(pkg)).toBe(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="${CORE_NAMESPACE}"><cp:revision>007</cp:revision></cp:coreProperties>`,
    );
    expect(partText(pkg).match(/xmlns:cp=/g)).toHaveLength(1);
    expect(readPresentationRevision(pkg)).toBe('007');

    const occupied = OpcPackage.create();
    const orphan = '<x:orphan xmlns:x="urn:test">KEEP</x:orphan>';
    occupied.setPart('/docProps/core.xml', orphan, 'application/xml');
    replaceRevision(occupied, '9');
    const allocated = occupied.relationships('/').find(({ type }) => type === CORE_RELATIONSHIP);
    expect(allocated?.resolvedTarget).toBe('/docProps/core1.xml');
    expect(partText(occupied)).toBe(orphan);
    expect(partText(occupied, '/docProps/core1.xml')).toContain(
      '<cp:revision>9</cp:revision>',
    );
    expect(partText(occupied, '/docProps/core1.xml')).not.toContain('creator');
  });

  it('losslessly inserts, replaces, preserves lexical values, and clears only revision', () => {
    const pkg = corePackage(coreXml(
      '\n  <d:title>Quarterly</d:title><d:subject>Forecast</d:subject><d:creator>Alice</d:creator>\n'
        + '  <c:lastModifiedBy>Editor</c:lastModifiedBy>\n'
        + '  <dcterms:created xsi:type="dcterms:W3CDTF">2026-07-30T00:00:00Z</dcterms:created>\n'
        + '  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-07-30T01:00:00Z</dcterms:modified>\n'
        + '  <!--KEEP--><x:opaque xmlns:x="urn:test">KEEP</x:opaque>\n',
      'c:coreProperties',
      `xmlns:c="${CORE_NAMESPACE}" xmlns:d="${DUBLIN_CORE_NAMESPACE}" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"`,
    ));
    pkg.setPart('/custom/keep.bin', new Uint8Array([1, 2, 3]), 'application/octet-stream');
    const unrelatedBefore = [...pkg.requirePart('/custom/keep.bin').bytes];

    replaceRevision(pkg, '007');
    let xml = partText(pkg);
    expect(xml).toContain('<c:revision>007</c:revision>');
    expect(readPresentationRevision(pkg)).toBe('007');
    expect(xml).toContain('<d:title>Quarterly</d:title>');
    expect(xml).toContain('<d:subject>Forecast</d:subject>');
    expect(xml).toContain('<d:creator>Alice</d:creator>');
    expect(xml).toContain('<c:lastModifiedBy>Editor</c:lastModifiedBy>');
    expect(xml).toContain('2026-07-30T00:00:00Z');
    expect(xml).toContain('2026-07-30T01:00:00Z');
    expect(xml).toContain('<!--KEEP-->');
    expect(xml).toContain('<x:opaque xmlns:x="urn:test">KEEP</x:opaque>');
    expect([...pkg.requirePart('/custom/keep.bin').bytes]).toEqual(unrelatedBefore);

    const beforeSame = packageSnapshot(pkg);
    replaceRevision(pkg, '007');
    expect(packageSnapshot(pkg)).toEqual(beforeSame);

    replaceRevision(pkg, '42');
    expect(readPresentationRevision(pkg)).toBe('42');
    expect(partText(pkg)).toContain('<c:revision>42</c:revision>');

    replaceRevision(pkg, '0009');
    expect(readPresentationRevision(pkg)).toBe('0009');
    expect(partText(pkg)).toContain('<c:revision>0009</c:revision>');

    replaceRevision(pkg, undefined);
    expect(readPresentationRevision(pkg)).toBeUndefined();
    xml = partText(pkg);
    expect(xml).not.toContain('<c:revision');
    expect(xml).toContain('<d:title>Quarterly</d:title>');
    expect(xml).toContain('<d:subject>Forecast</d:subject>');
    expect(xml).toContain('<d:creator>Alice</d:creator>');
    expect(xml).toContain('<c:lastModifiedBy>Editor</c:lastModifiedBy>');
    expect(pkg.relationships('/').filter(({ type }) => type === CORE_RELATIONSHIP)).toHaveLength(1);

    const beforeAbsentClear = packageSnapshot(pkg);
    replaceRevision(pkg, undefined);
    expect(packageSnapshot(pkg)).toEqual(beforeAbsentClear);
  });

  it('repairs simple invalid and self-closing revision state while preserving unsafe XML', () => {
    const invalid = corePackage(coreXml(
      '<cp:revision>abc</cp:revision><dc:title>KEEP</dc:title>',
    ));
    expect(readPresentationRevision(invalid)).toBeUndefined();
    replaceRevision(invalid, '7');
    expect(readPresentationRevision(invalid)).toBe('7');
    expect(partText(invalid)).toContain('<cp:revision>7</cp:revision>');
    expect(partText(invalid)).toContain('<dc:title>KEEP</dc:title>');

    const clearInvalid = corePackage(coreXml('<cp:revision>-1</cp:revision>'));
    replaceRevision(clearInvalid, undefined);
    expect(partText(clearInvalid)).not.toContain('<cp:revision');

    const selfClosingRoot = corePackage(
      `<?xml version="1.0"?><c:coreProperties xmlns:c="${CORE_NAMESPACE}"/>`,
    );
    replaceRevision(selfClosingRoot, '8');
    expect(partText(selfClosingRoot)).toContain('<c:revision>8</c:revision>');

    const selfClosingRevision = corePackage(coreXml(
      '<cp:revision custom="KEEP"/><dc:title>KEEP</dc:title>',
    ));
    expect(readPresentationRevision(selfClosingRevision)).toBeUndefined();
    replaceRevision(selfClosingRevision, '9');
    expect(partText(selfClosingRevision)).toContain(
      '<cp:revision custom="KEEP">9</cp:revision>',
    );
    expect(partText(selfClosingRevision)).toContain('<dc:title>KEEP</dc:title>');

    const wrongNamespace = corePackage(coreXml(
      '<x:revision xmlns:x="urn:wrong">KEEP</x:revision>',
      'cp:coreProperties',
      `xmlns:cp="${CORE_NAMESPACE}"`,
    ));
    replaceRevision(wrongNamespace, '10');
    expect(partText(wrongNamespace)).toContain(
      '<x:revision xmlns:x="urn:wrong">KEEP</x:revision>',
    );
    expect(partText(wrongNamespace)).toContain('<cp:revision>10</cp:revision>');
  });

  it('rejects invalid inputs and unsafe ownership without mutation', () => {
    const invalidValues = [
      '',
      ' ',
      '+1',
      '-1',
      '1.0',
      '1e3',
      '１２',
      null,
      false,
      0,
      1n,
      {},
      [],
      Symbol('revision'),
    ] as const;
    for (const value of invalidValues) {
      const pkg = OpcPackage.create();
      const before = packageSnapshot(pkg);
      expect(() => replaceRevision(pkg, value)).toThrow(TypeError);
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    const unsafe: OpcPackage[] = [
      corePackage(coreXml(
        '<cp:revision>1</cp:revision><cp:revision>2</cp:revision>',
      )),
      corePackage(coreXml(
        '<cp:revision>1<x:keep xmlns:x="urn:test"/>2</cp:revision>',
      )),
      corePackage(coreXml('<cp:revision><![CDATA[7]]></cp:revision>')),
      corePackage('<x:notCore xmlns:x="urn:test"/>'),
      corePackage(coreXml('<cp:revision>7</cp:revision>'), {
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

    const duplicate = corePackage(coreXml('<cp:revision>1</cp:revision>'));
    duplicate.transaction(() => {
      duplicate.setPart(
        '/docProps/core2.xml',
        coreXml('<cp:revision>2</cp:revision>'),
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
      expect(() => replaceRevision(pkg, '42')).toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
    }
  });

  it('rolls back minimal creation and existing revision edits', () => {
    const created = OpcPackage.create();
    const createdBefore = packageSnapshot(created);

    expect(() => created.transaction(() => {
      replacePresentationRevision(created, '7');
      expect(readPresentationRevision(created)).toBe('7');
      throw new Error('restore revision creation');
    })).toThrow('restore revision creation');
    expect(packageSnapshot(created)).toEqual(createdBefore);
    expect(created.hasPart('/docProps/core.xml')).toBe(false);
    expect(created.relationships('/')).toHaveLength(0);

    const existing = corePackage(coreXml(
      '<cp:revision>1</cp:revision><dc:creator>Alice</dc:creator>',
    ));
    const existingBefore = packageSnapshot(existing);
    expect(() => existing.transaction(() => {
      replacePresentationRevision(existing, '99');
      expect(readPresentationRevision(existing)).toBe('99');
      throw new Error('restore revision edit');
    })).toThrow('restore revision edit');
    expect(packageSnapshot(existing)).toEqual(existingBefore);
    expect(readPresentationRevision(existing)).toBe('1');
    expect(partText(existing)).toContain('<dc:creator>Alice</dc:creator>');
  });
});
