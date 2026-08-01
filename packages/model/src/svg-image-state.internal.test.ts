import { describe, expect, it } from 'vitest';
import { LosslessXmlDocument } from '@pptx/lossless-xml';
import { OpcPackage, type Relationship } from '@pptx/opc';
import {
  attributeNamespaceUri,
  elementNamespaceUri,
  hasSvgImageExtensionCandidate,
  readSvgImageState,
  relationshipReferenceCount,
} from './svg-image-state.internal.js';

const PRESENTATION_NAMESPACE =
  'http://schemas.openxmlformats.org/presentationml/2006/main';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SVG_NAMESPACE =
  'http://schemas.microsoft.com/office/drawing/2016/SVG/main';
const SVG_EXTENSION_URI = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';
const IMAGE_RELATIONSHIP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

describe('SVG image namespace resolution', () => {
  it('resolves inherited default, element, and attribute prefixes independently', () => {
    const xml = LosslessXmlDocument.parse(
      `<root xmlns="urn:default" xmlns:d="${DRAWING_NAMESPACE}" `
        + `xmlns:rel="${RELATIONSHIP_NAMESPACE}">`
        + '<d:blip rel:embed="rId1" plain="value"><child/></d:blip></root>',
    );
    const root = xml.roots[0]!;
    const blip = xml.elements('blip')[0]!;
    const child = xml.elements('child')[0]!;

    expect(elementNamespaceUri(root)).toBe('urn:default');
    expect(elementNamespaceUri(blip)).toBe(DRAWING_NAMESPACE);
    expect(elementNamespaceUri(child)).toBe('urn:default');
    expect(attributeNamespaceUri(blip, xml.attribute(blip, 'rel:embed')!))
      .toBe(RELATIONSHIP_NAMESPACE);
    expect(attributeNamespaceUri(blip, xml.attribute(blip, 'plain')!)).toBeUndefined();
    expect(attributeNamespaceUri(root, xml.attribute(root, 'xmlns')!))
      .toBe('http://www.w3.org/2000/xmlns/');
  });
});

describe('SVG image paired state', () => {
  it('reads canonical PptxGenJS state without mutating XML or package state', () => {
    const fixture = svgFixture(canonicalBlip());
    const beforeXml = fixture.xml.serialize();
    const beforeJournal = [...fixture.pkg.mutations];
    const state = readSvgImageState(
      fixture.xml,
      fixture.picture,
      fixture.relationships,
      fixture.pkg,
    );

    expect(state).toMatchObject({
      fallbackPartUri: '/ppt/media/fallback.png',
      svgPartUri: '/ppt/media/vector.svg',
      fallbackRelationship: {
        id: 'rId1',
        type: IMAGE_RELATIONSHIP,
        targetMode: 'Internal',
      },
      svgRelationship: {
        id: 'rId2',
        type: IMAGE_RELATIONSHIP,
        targetMode: 'Internal',
      },
    });
    expect(state?.fallbackReference.value).toBe('rId1');
    expect(state?.svgReference.value).toBe('rId2');
    expect(hasSvgImageExtensionCandidate(fixture.xml, fixture.picture)).toBe(true);
    expect(relationshipReferenceCount(fixture.xml, 'rId1')).toBe(1);
    expect(relationshipReferenceCount(fixture.xml, 'rId2')).toBe(1);
    expect(fixture.xml.serialize()).toBe(beforeXml);
    expect(fixture.pkg.mutations).toEqual(beforeJournal);
  });

  it('reads alternate prefixes, inherited bindings, and LibreOffice image/svg parts', () => {
    const fixture = svgFixture(
      '<d:blip rel:embed="rId1"><d:extLst>'
        + `<d:ext uri="${SVG_EXTENSION_URI}">`
        + '<v:svgBlip rel:embed="rId2"/>'
        + '</d:ext></d:extLst></d:blip>',
      {
        rootNamespaces: `xmlns:d="${DRAWING_NAMESPACE}" `
          + `xmlns:rel="${RELATIONSHIP_NAMESPACE}" xmlns:v="${SVG_NAMESPACE}"`,
        svgContentType: 'image/svg',
      },
    );
    const state = readSvgImageState(
      fixture.xml,
      fixture.picture,
      fixture.relationships,
      fixture.pkg,
    );

    expect(state?.fallbackPartUri).toBe('/ppt/media/fallback.png');
    expect(state?.svgPartUri).toBe('/ppt/media/vector.svg');
    expect(state?.fallbackReference.name).toBe('rel:embed');
    expect(state?.svgReference.name).toBe('rel:embed');
    expect(fixture.pkg.requirePart(state!.svgPartUri).contentType).toBe('image/svg');
    expect(relationshipReferenceCount(fixture.xml, 'rId1')).toBe(1);
  });

  it('counts relationship references by namespace rather than lexical prefix', () => {
    const fixture = svgFixture(
      '<d:blip rel:embed="rId1"><d:extLst>'
        + `<d:ext uri="${SVG_EXTENSION_URI}">`
        + '<v:svgBlip rel:embed="rId2"/><x:keep xmlns:x="urn:keep" x:id="rId1"/>'
        + '</d:ext></d:extLst></d:blip>',
      {
        rootNamespaces: `xmlns:d="${DRAWING_NAMESPACE}" `
          + `xmlns:rel="${RELATIONSHIP_NAMESPACE}" xmlns:v="${SVG_NAMESPACE}"`,
      },
    );
    const source = fixture.xml.source.replace(
      '</p:pic>',
      '<x:other xmlns:x="urn:other" xmlns:q="'
        + RELATIONSHIP_NAMESPACE
        + '" q:embed="rId1"/></p:pic>',
    );
    const xml = LosslessXmlDocument.parse(source);

    expect(relationshipReferenceCount(xml, 'rId1')).toBe(2);
    expect(relationshipReferenceCount(xml, 'rId2')).toBe(1);
  });

  it('rejects malformed or namespace-confused SVG structures while retaining candidates', () => {
    const malformed = [
      '<a:blip r:embed="rId1"/>',
      '<a:blip r:embed="rId1"><a:extLst>'
        + `<a:ext uri="${SVG_EXTENSION_URI}"><x:svgBlip xmlns:x="urn:wrong" `
        + 'r:embed="rId2"/></a:ext></a:extLst></a:blip>',
      '<a:blip embed="rId1"><a:extLst>'
        + `<a:ext uri="${SVG_EXTENSION_URI}"><asvg:svgBlip xmlns:asvg="${SVG_NAMESPACE}" `
        + 'r:embed="rId2"/></a:ext></a:extLst></a:blip>',
      '<a:blip r:embed="rId1"><x:extLst xmlns:x="urn:wrong">'
        + `<x:ext uri="${SVG_EXTENSION_URI}"><asvg:svgBlip xmlns:asvg="${SVG_NAMESPACE}" `
        + 'r:embed="rId2"/></x:ext></x:extLst></a:blip>',
      '<a:blip r:embed="rId1"><a:extLst>'
        + `<a:ext uri="${SVG_EXTENSION_URI}"><asvg:svgBlip xmlns:asvg="${SVG_NAMESPACE}" `
        + 'r:embed="rId2"/></a:ext>'
        + `<a:ext uri="${SVG_EXTENSION_URI}"><asvg:svgBlip xmlns:asvg="${SVG_NAMESPACE}" `
        + 'r:embed="rId2"/></a:ext></a:extLst></a:blip>',
      '<a:blip r:embed="rId1"><a:extLst>'
        + `<a:ext uri="${SVG_EXTENSION_URI}"><a:wrapper>`
        + `<asvg:svgBlip xmlns:asvg="${SVG_NAMESPACE}" r:embed="rId2"/>`
        + '</a:wrapper></a:ext></a:extLst></a:blip>',
      '<a:blip r:embed="rId1"><a:extLst>'
        + `<a:notExt uri="${SVG_EXTENSION_URI}">`
        + `<asvg:svgBlip xmlns:asvg="${SVG_NAMESPACE}" r:embed="rId2"/>`
        + '</a:notExt></a:extLst></a:blip>',
      '<a:blip r:embed="rId1"><a:extLst>'
        + `<a:ext uri="${SVG_EXTENSION_URI}"><asvg:svgBlip xmlns:asvg="${SVG_NAMESPACE}" `
        + 'r:embed="rId2"/><x:extra xmlns:x="urn:extra"/></a:ext></a:extLst></a:blip>',
    ];

    for (const blip of malformed) {
      const fixture = svgFixture(blip);
      const beforeXml = fixture.xml.serialize();
      const beforeJournal = [...fixture.pkg.mutations];
      expect(readSvgImageState(
        fixture.xml,
        fixture.picture,
        fixture.relationships,
        fixture.pkg,
      ), blip).toBeUndefined();
      expect(hasSvgImageExtensionCandidate(fixture.xml, fixture.picture), blip)
        .toBe(blip !== '<a:blip r:embed="rId1"/>');
      expect(fixture.xml.serialize()).toBe(beforeXml);
      expect(fixture.pkg.mutations).toEqual(beforeJournal);
    }
  });

  it('rejects unsafe relationship state and missing targets', () => {
    const fixture = svgFixture(canonicalBlip());
    const base = fixture.relationships;
    const cases: readonly Relationship[][] = [
      base.filter(({ id }) => id !== 'rId2'),
      base.map((relationship) => relationship.id === 'rId2'
        ? { ...relationship, type: 'urn:wrong' }
        : relationship),
      base.map((relationship) => relationship.id === 'rId2'
        ? {
            id: relationship.id,
            type: relationship.type,
            target: 'https://example.com/vector.svg',
            targetMode: 'External' as const,
          }
        : relationship),
      base.map((relationship) => relationship.id === 'rId2'
        ? { ...relationship, resolvedTarget: '/ppt/media/missing.svg' }
        : relationship),
      base.map((relationship) => relationship.id === 'rId2'
        ? { ...relationship, resolvedTarget: '/ppt/media/fallback.png' }
        : relationship),
      [...base, { ...base[1]! }],
    ];

    for (const relationships of cases) {
      expect(readSvgImageState(
        fixture.xml,
        fixture.picture,
        relationships,
        fixture.pkg,
      )).toBeUndefined();
    }
  });
});

interface SvgFixtureOptions {
  readonly rootNamespaces?: string;
  readonly svgContentType?: string;
}

function svgFixture(blip: string, options: SvgFixtureOptions = {}) {
  const rootNamespaces = options.rootNamespaces ?? '';
  const source = `<p:sld xmlns:p="${PRESENTATION_NAMESPACE}" `
    + `xmlns:a="${DRAWING_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}" `
    + `${rootNamespaces}><p:cSld><p:spTree><p:pic><p:nvPicPr>`
    + '<p:cNvPr id="2" name="SVG"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
    + `<p:blipFill>${blip}<a:stretch/></p:blipFill><p:spPr/></p:pic>`
    + '</p:spTree></p:cSld></p:sld>';
  const pkg = OpcPackage.create();
  pkg.setPart(
    '/ppt/slides/slide1.xml',
    source,
    'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  );
  pkg.setPart('/ppt/media/fallback.png', new Uint8Array([1]), 'image/png');
  pkg.setPart(
    '/ppt/media/vector.svg',
    new Uint8Array([2]),
    options.svgContentType ?? 'image/svg+xml',
  );
  pkg.addRelationship('/ppt/slides/slide1.xml', {
    id: 'rId1',
    type: IMAGE_RELATIONSHIP,
    target: '../media/fallback.png',
    targetMode: 'Internal',
  });
  pkg.addRelationship('/ppt/slides/slide1.xml', {
    id: 'rId2',
    type: IMAGE_RELATIONSHIP,
    target: '../media/vector.svg',
    targetMode: 'Internal',
  });
  const xml = LosslessXmlDocument.parse(source);
  const picture = xml.elements('pic')[0];
  if (!picture) throw new Error('Fixture has no picture');
  return {
    pkg,
    xml,
    picture,
    relationships: pkg.relationships('/ppt/slides/slide1.xml'),
  };
}

function canonicalBlip(): string {
  return '<a:blip r:embed="rId1"><a:extLst>'
    + `<a:ext uri="${SVG_EXTENSION_URI}">`
    + `<asvg:svgBlip xmlns:asvg="${SVG_NAMESPACE}" r:embed="rId2"/>`
    + '</a:ext></a:extLst></a:blip>';
}
