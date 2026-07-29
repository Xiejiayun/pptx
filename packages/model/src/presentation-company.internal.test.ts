import { describe, expect, it } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import {
  readPresentationCompany,
  replacePresentationCompany,
} from './presentation-company.internal.js';

const APP_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';
const APP_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.extended-properties+xml';
const RELATIONSHIPS_CONTENT_TYPE =
  'application/vnd.openxmlformats-package.relationships+xml';
const APP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
const VT_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';

const appXml = (
  children: string,
  root = 'Properties',
  namespaces = `xmlns="${APP_NAMESPACE}" xmlns:vt="${VT_NAMESPACE}"`,
): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><${root} ${namespaces}>${children}</${root}>`;

function appPackage(
  xml?: string,
  options: { readonly uri?: string; readonly contentType?: string } = {},
): OpcPackage {
  const pkg = OpcPackage.create();
  if (xml !== undefined) {
    const uri = options.uri ?? '/docProps/app.xml';
    pkg.transaction(() => {
      pkg.setPart(uri, xml, options.contentType ?? APP_CONTENT_TYPE);
      pkg.addRelationship('/', {
        id: 'rId1',
        type: APP_RELATIONSHIP,
        target: uri.slice(1),
      });
    });
  }
  return pkg;
}

function replaceCompany(pkg: OpcPackage, value: unknown): void {
  pkg.transaction(() => {
    replacePresentationCompany(pkg, value as never);
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

function partText(pkg: OpcPackage, uri = '/docProps/app.xml'): string {
  return new TextDecoder().decode(pkg.requirePart(uri).bytes);
}

function setRootRelationships(pkg: OpcPackage, relationships: string): void {
  pkg.setPart(
    '/_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`,
    RELATIONSHIPS_CONTENT_TYPE,
  );
}

describe('presentation company extended property', () => {
  it('reads only one namespace-correct direct simple Company without mutation', () => {
    const cases: readonly [OpcPackage, string | undefined][] = [
      [appPackage(), undefined],
      [appPackage(appXml('<Application>PowerPoint</Application>')), undefined],
      [appPackage(appXml(
        '<ep:Company>作者 Acme &amp; Partners</ep:Company><ep:AppVersion>16.0</ep:AppVersion>',
        'ep:Properties',
        `xmlns:ep="${APP_NAMESPACE}"`,
      ), { uri: '/metadata/application.xml' }), '作者 Acme & Partners'],
      [appPackage(appXml('<Company/>')), ''],
      [appPackage(appXml('<Company xmlns="urn:wrong">Wrong</Company>')), undefined],
      [appPackage(appXml('<Opaque><Company>Nested</Company></Opaque>')), undefined],
      [appPackage(appXml('<Company>One</Company><Company>Two</Company>')), undefined],
      [appPackage(appXml('<Company>Before<Opaque/>After</Company>')), undefined],
      [appPackage(appXml('<Company><![CDATA[Raw]]></Company>')), undefined],
      [appPackage(
        `<x:notProperties xmlns:x="urn:test"><Company xmlns="${APP_NAMESPACE}">Wrong root</Company></x:notProperties>`,
      ), undefined],
      [appPackage(appXml('<Company>Wrong type</Company>'), {
        contentType: 'application/xml',
      }), undefined],
    ];

    for (const [pkg, expected] of cases) {
      const before = packageSnapshot(pkg);
      expect(readPresentationCompany(pkg)).toBe(expected);
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    const dangling = OpcPackage.create();
    setRootRelationships(
      dangling,
      `<Relationship Id="rId1" Type="${APP_RELATIONSHIP}" Target="docProps/missing.xml"/>`,
    );
    const danglingBefore = packageSnapshot(dangling);
    expect(readPresentationCompany(dangling)).toBeUndefined();
    expect(packageSnapshot(dangling)).toEqual(danglingBefore);

    const external = OpcPackage.create();
    setRootRelationships(
      external,
      `<Relationship Id="rId1" Type="${APP_RELATIONSHIP}" Target="https://example.com/app.xml" TargetMode="External"/>`,
    );
    const externalBefore = packageSnapshot(external);
    expect(readPresentationCompany(external)).toBeUndefined();
    expect(packageSnapshot(external)).toEqual(externalBefore);

    const duplicate = appPackage(appXml('<Company>One</Company>'));
    duplicate.transaction(() => {
      duplicate.setPart(
        '/docProps/app2.xml',
        appXml('<Company>Two</Company>'),
        APP_CONTENT_TYPE,
      );
      duplicate.addRelationship('/', {
        id: 'rId2',
        type: APP_RELATIONSHIP,
        target: 'docProps/app2.xml',
      });
    });
    const duplicateBefore = packageSnapshot(duplicate);
    expect(readPresentationCompany(duplicate)).toBeUndefined();
    expect(packageSnapshot(duplicate)).toEqual(duplicateBefore);
  });

  it('creates only a minimal Company part and avoids an occupied canonical URI', () => {
    const pkg = OpcPackage.create();
    replaceCompany(pkg, '作者 Acme & <Partners>');

    const relationship = pkg.relationships('/').find(({ type }) => type === APP_RELATIONSHIP);
    expect(relationship).toMatchObject({
      target: 'docProps/app.xml',
      targetMode: 'Internal',
      resolvedTarget: '/docProps/app.xml',
    });
    expect(pkg.requirePart('/docProps/app.xml').contentType).toBe(APP_CONTENT_TYPE);
    expect(partText(pkg)).toBe(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="${APP_NAMESPACE}"><Company>作者 Acme &amp; &lt;Partners&gt;</Company></Properties>`,
    );
    expect(readPresentationCompany(pkg)).toBe('作者 Acme & <Partners>');
    expect(partText(pkg)).not.toContain('Application');
    expect(partText(pkg)).not.toContain('AppVersion');

    const occupied = OpcPackage.create();
    const orphan = '<x:orphan xmlns:x="urn:test">KEEP</x:orphan>';
    occupied.setPart('/docProps/app.xml', orphan, 'application/xml');
    replaceCompany(occupied, 'Allocated');
    const allocated = occupied.relationships('/').find(({ type }) => type === APP_RELATIONSHIP);
    expect(allocated?.resolvedTarget).toBe('/docProps/app1.xml');
    expect(partText(occupied)).toBe(orphan);
    expect(readPresentationCompany(occupied)).toBe('Allocated');
    expect(partText(occupied, '/docProps/app1.xml')).not.toContain('Application');
  });

  it('inserts before following properties with the root lexical namespace form', () => {
    const defaultNamespace = appPackage(appXml(
      '<Application>PowerPoint</Application><LinksUpToDate>false</LinksUpToDate>'
        + '<AppVersion>16.0</AppVersion>',
    ));
    replaceCompany(defaultNamespace, 'Default');
    const defaultXml = partText(defaultNamespace);
    expect(defaultXml).toContain('<Company>Default</Company>');
    expect(defaultXml.indexOf('<Company>')).toBeLessThan(defaultXml.indexOf('<LinksUpToDate>'));

    const prefixed = appPackage(appXml(
      '<ep:Application>PowerPoint</ep:Application><ep:AppVersion>16.0</ep:AppVersion>',
      'ep:Properties',
      `xmlns:ep="${APP_NAMESPACE}"`,
    ));
    replaceCompany(prefixed, 'Prefixed');
    const prefixedXml = partText(prefixed);
    expect(prefixedXml).toContain('<ep:Company>Prefixed</ep:Company>');
    expect(prefixedXml.indexOf('<ep:Company>')).toBeLessThan(
      prefixedXml.indexOf('<ep:AppVersion>'),
    );
  });

  it('losslessly replaces, empties, and clears only Company', () => {
    const pkg = appPackage(appXml(
      '\n  <Application>Microsoft Office PowerPoint</Application>\n'
        + '  <PresentationFormat>Custom</PresentationFormat><Slides>4</Slides>\n'
        + '  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant>'
        + '<vt:lpstr>Theme</vt:lpstr></vt:variant></vt:vector></HeadingPairs>\n'
        + '  <Company custom="KEEP">Original</Company>\n'
        + '  <LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>\n'
        + '  <!--KEEP--><Opaque xmlns="urn:test">KEEP</Opaque>\n'
        + '  <AppVersion>16.0000</AppVersion>\n',
    ));
    pkg.setPart('/custom/keep.bin', new Uint8Array([1, 2, 3]), 'application/octet-stream');
    const unrelatedBefore = [...pkg.requirePart('/custom/keep.bin').bytes];

    replaceCompany(pkg, 'Replacement & <safe>');
    let xml = partText(pkg);
    expect(xml).toContain(
      '<Company custom="KEEP">Replacement &amp; &lt;safe&gt;</Company>',
    );
    expect(xml).toContain('<Application>Microsoft Office PowerPoint</Application>');
    expect(xml).toContain('<PresentationFormat>Custom</PresentationFormat><Slides>4</Slides>');
    expect(xml).toContain('<vt:lpstr>Theme</vt:lpstr>');
    expect(xml).toContain('<LinksUpToDate>false</LinksUpToDate>');
    expect(xml).toContain('<SharedDoc>false</SharedDoc>');
    expect(xml).toContain('<AppVersion>16.0000</AppVersion>');
    expect(xml).toContain('<!--KEEP-->');
    expect(xml).toContain('<Opaque xmlns="urn:test">KEEP</Opaque>');
    expect([...pkg.requirePart('/custom/keep.bin').bytes]).toEqual(unrelatedBefore);

    const beforeSame = packageSnapshot(pkg);
    replaceCompany(pkg, 'Replacement & <safe>');
    expect(packageSnapshot(pkg)).toEqual(beforeSame);

    replaceCompany(pkg, 'Next');
    expect(readPresentationCompany(pkg)).toBe('Next');
    expect(partText(pkg)).toContain('<Company custom="KEEP">Next</Company>');

    replaceCompany(pkg, '');
    expect(readPresentationCompany(pkg)).toBe('');
    expect(partText(pkg)).toContain('<Company custom="KEEP"></Company>');

    replaceCompany(pkg, undefined);
    expect(readPresentationCompany(pkg)).toBeUndefined();
    xml = partText(pkg);
    expect(xml).not.toContain('<Company');
    expect(xml).toContain('<Application>Microsoft Office PowerPoint</Application>');
    expect(xml).toContain('<AppVersion>16.0000</AppVersion>');
    expect(pkg.relationships('/').filter(({ type }) => type === APP_RELATIONSHIP)).toHaveLength(1);

    const beforeAbsentClear = packageSnapshot(pkg);
    replaceCompany(pkg, undefined);
    expect(packageSnapshot(pkg)).toEqual(beforeAbsentClear);
  });

  it('preserves semantic no-ops and safely expands self-closing state', () => {
    const numericEntity = appPackage(appXml(
      '<Company>A&#38;B</Company><AppVersion>16.0</AppVersion>',
    ));
    const beforeEntityNoOp = packageSnapshot(numericEntity);
    replaceCompany(numericEntity, 'A&B');
    expect(packageSnapshot(numericEntity)).toEqual(beforeEntityNoOp);

    const selfClosingRoot = appPackage(
      `<?xml version="1.0"?><ep:Properties xmlns:ep="${APP_NAMESPACE}"/>`,
    );
    replaceCompany(selfClosingRoot, 'Root expanded');
    expect(readPresentationCompany(selfClosingRoot)).toBe('Root expanded');
    expect(partText(selfClosingRoot)).toContain(
      '<ep:Company>Root expanded</ep:Company>',
    );

    const selfClosingCompany = appPackage(appXml(
      '<Company custom="KEEP"/><AppVersion>16.0</AppVersion>',
    ));
    replaceCompany(selfClosingCompany, 'Expanded');
    expect(partText(selfClosingCompany)).toContain(
      '<Company custom="KEEP">Expanded</Company>',
    );
    expect(partText(selfClosingCompany)).toContain('<AppVersion>16.0</AppVersion>');

    const wrongNamespace = appPackage(appXml(
      '<x:Company xmlns:x="urn:wrong">KEEP</x:Company><AppVersion>16.0</AppVersion>',
    ));
    replaceCompany(wrongNamespace, 'Direct');
    expect(partText(wrongNamespace)).toContain(
      '<x:Company xmlns:x="urn:wrong">KEEP</x:Company>',
    );
    expect(partText(wrongNamespace)).toContain('<Company>Direct</Company>');
  });

  it('rejects invalid values and unsafe ownership without mutation', () => {
    for (const value of [
      null,
      true,
      false,
      0,
      1,
      {},
      [],
      Symbol('company'),
      'bad\u0001company',
    ]) {
      const pkg = OpcPackage.create();
      const before = packageSnapshot(pkg);
      expect(() => replaceCompany(pkg, value)).toThrow(TypeError);
      expect(packageSnapshot(pkg)).toEqual(before);
    }

    const unsafe: OpcPackage[] = [
      appPackage(appXml('<Company>One</Company><Company>Two</Company>')),
      appPackage(appXml('<Company>Before<Opaque/>After</Company>')),
      appPackage(appXml('<Company><![CDATA[Raw]]></Company>')),
      appPackage('<x:notProperties xmlns:x="urn:test"/>'),
      appPackage(appXml('<Company>Wrong type</Company>'), {
        contentType: 'application/xml',
      }),
    ];

    const dangling = OpcPackage.create();
    setRootRelationships(
      dangling,
      `<Relationship Id="rId1" Type="${APP_RELATIONSHIP}" Target="docProps/missing.xml"/>`,
    );
    unsafe.push(dangling);

    const external = OpcPackage.create();
    setRootRelationships(
      external,
      `<Relationship Id="rId1" Type="${APP_RELATIONSHIP}" Target="https://example.com/app.xml" TargetMode="External"/>`,
    );
    unsafe.push(external);

    const duplicate = appPackage(appXml('<Company>One</Company>'));
    duplicate.transaction(() => {
      duplicate.setPart(
        '/docProps/app2.xml',
        appXml('<Company>Two</Company>'),
        APP_CONTENT_TYPE,
      );
      duplicate.addRelationship('/', {
        id: 'rId2',
        type: APP_RELATIONSHIP,
        target: 'docProps/app2.xml',
      });
    });
    unsafe.push(duplicate);

    for (const pkg of unsafe) {
      const before = packageSnapshot(pkg);
      expect(() => replaceCompany(pkg, 'Replacement')).toThrow();
      expect(packageSnapshot(pkg)).toEqual(before);
    }
  });

  it('rolls back minimal creation and existing Company edits', () => {
    const created = OpcPackage.create();
    const createdBefore = packageSnapshot(created);

    expect(() => created.transaction(() => {
      replacePresentationCompany(created, 'Temporary');
      expect(readPresentationCompany(created)).toBe('Temporary');
      throw new Error('restore company creation');
    })).toThrow('restore company creation');

    expect(packageSnapshot(created)).toEqual(createdBefore);
    expect(created.hasPart('/docProps/app.xml')).toBe(false);
    expect(created.relationships('/')).toHaveLength(0);

    const existing = appPackage(appXml(
      '<Company>Original</Company><Application>PowerPoint</Application>',
    ));
    const existingBefore = packageSnapshot(existing);
    expect(() => existing.transaction(() => {
      replacePresentationCompany(existing, 'Temporary');
      expect(readPresentationCompany(existing)).toBe('Temporary');
      throw new Error('restore company edit');
    })).toThrow('restore company edit');
    expect(packageSnapshot(existing)).toEqual(existingBefore);
    expect(readPresentationCompany(existing)).toBe('Original');
    expect(partText(existing)).toContain('<Application>PowerPoint</Application>');
  });
});
