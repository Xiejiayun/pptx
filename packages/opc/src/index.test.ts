import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import {
  joinPartUri,
  normalizePartUri,
  OpcPackage,
  partUriBasename,
  partUriDirname,
  partUriExtension,
  relativeRelationshipTarget,
  relationshipPartUri,
  resolveRelationshipTarget,
  sourcePartUri,
} from './index.js';

async function fixture(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><x:Unknown xmlns:x="urn:test" keep="yes"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="officeDocument" Target="ppt/presentation.xml"/></Relationships>');
  zip.file('ppt/presentation.xml', '<p:presentation xmlns:p="p"><p:sldIdLst><p:sldId id="256" r:id="rId1" xmlns:r="r"/></p:sldIdLst></p:presentation>');
  zip.file('ppt/_rels/presentation.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="slide" Target="slides/slide1.xml"/></Relationships>');
  zip.file('ppt/slides/slide1.xml', '<p:sld xmlns:p="p"><x:unknown xmlns:x="x" keep="true"/></p:sld>');
  zip.file('docProps/opaque.bin', new Uint8Array([0, 1, 2, 3, 255]));
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

describe('OpcPackage', () => {
  it('returns the original bytes if there are no mutations', async () => {
    const input = await fixture();
    const pkg = await OpcPackage.open(input);
    expect(await pkg.write()).toEqual(input);
  });

  it('builds resolved relationships and preserves untouched payloads', async () => {
    const input = await fixture();
    const pkg = await OpcPackage.open(input);
    expect(pkg.relationships('/ppt/presentation.xml')[0]?.resolvedTarget).toBe('/ppt/slides/slide1.xml');
    const opaqueHash = hash(pkg.requirePart('/docProps/opaque.bin').bytes);
    pkg.setPart('/ppt/slides/slide1.xml', '<p:sld xmlns:p="p"><p:changed/></p:sld>');
    const reopened = await OpcPackage.open(await pkg.write());
    expect(hash(reopened.requirePart('/docProps/opaque.bin').bytes)).toBe(opaqueHash);
  });

  it('resolves relationship targets relative to the source part', () => {
    expect(resolveRelationshipTarget('/ppt/slides/slide1.xml', '../media/image1.png')).toBe('/ppt/media/image1.png');
  });

  it('handles OPC part URIs without host path semantics', () => {
    expect(normalizePartUri('ppt//slides/./../media/image1.png')).toBe('/ppt/media/image1.png');
    expect(partUriDirname('/ppt/slides/slide1.xml')).toBe('/ppt/slides');
    expect(partUriBasename('/ppt/slides/slide1.xml')).toBe('slide1.xml');
    expect(partUriBasename('/_rels/.rels', '.rels')).toBe('');
    expect(partUriExtension('/ppt/slides/slide1.xml')).toBe('.xml');
    expect(partUriExtension('/_rels/.rels')).toBe('');
    expect(joinPartUri('/ppt/slideMasters', '../slideLayouts', 'slideLayout1.xml')).toBe(
      '/ppt/slideLayouts/slideLayout1.xml',
    );
    expect(relativeRelationshipTarget('/ppt/slides/slide1.xml', '/ppt/media/image1.png')).toBe(
      '../media/image1.png',
    );
    expect(relativeRelationshipTarget('/', '/ppt/presentation.xml')).toBe('ppt/presentation.xml');
    expect(relationshipPartUri('/')).toBe('/_rels/.rels');
    expect(sourcePartUri('/_rels/.rels')).toBe('/');
    expect(relationshipPartUri('/ppt/slides/slide1.xml')).toBe('/ppt/slides/_rels/slide1.xml.rels');
    expect(sourcePartUri('/ppt/slides/_rels/slide1.xml.rels')).toBe('/ppt/slides/slide1.xml');
    expect(() => normalizePartUri('../../outside.xml')).toThrow(/Invalid part URI/);
    expect(() => normalizePartUri('/ppt\\outside.xml')).toThrow(/Invalid part URI/);
  });

  it('adds and removes parts, content types, and relationships together', async () => {
    const pkg = await OpcPackage.open(await fixture());
    pkg.setPart('/ppt/media/image1.png', new Uint8Array([137, 80, 78, 71]), 'image/png');
    const relationship = pkg.addRelationship('/ppt/slides/slide1.xml', {
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      target: '../media/image1.png',
    });
    expect(relationship.id).toBe('rId1');
    expect(pkg.graph.find(({ uri }) => uri === '/ppt/media/image1.png')?.incoming[0]?.sourceUri).toBe(
      '/ppt/slides/slide1.xml',
    );

    const reopened = await OpcPackage.open(await pkg.write());
    expect(reopened.requirePart('/ppt/media/image1.png').contentType).toBe('image/png');
    expect(reopened.relationships('/ppt/slides/slide1.xml')[0]?.resolvedTarget).toBe('/ppt/media/image1.png');
    expect(new TextDecoder().decode(reopened.requirePart('/[Content_Types].xml').bytes)).toContain(
      '<x:Unknown xmlns:x="urn:test" keep="yes"/>',
    );
    reopened.setPart('/ppt/media/image1.png', new Uint8Array([1, 2, 3]), 'image/jpeg');
    const retyped = await OpcPackage.open(await reopened.write());
    expect(retyped.requirePart('/ppt/media/image1.png').contentType).toBe('image/jpeg');
    const external = reopened.updateRelationship('/ppt/slides/slide1.xml', relationship.id, {
      target: 'https://example.com/image.png',
      targetMode: 'External',
    });
    expect(external.resolvedTarget).toBeUndefined();
    expect(external.targetMode).toBe('External');
    expect(reopened.removeRelationship('/ppt/slides/slide1.xml', relationship.id)).toBe(true);

    reopened.deletePart('/ppt/media/image1.png');
    const deleted = await OpcPackage.open(await reopened.write());
    expect(deleted.hasPart('/ppt/media/image1.png')).toBe(false);
    expect(deleted.relationships('/ppt/slides/slide1.xml')).toHaveLength(0);
  });

  it('enforces entry and part resource budgets', async () => {
    const input = await fixture();
    await expect(OpcPackage.open(input, { limits: { maxEntries: 2 } })).rejects.toThrow(/entries/);
    await expect(OpcPackage.open(input, { limits: { maxPartBytes: 3 } })).rejects.toThrow(/exceeds/);
  });

  it('rejects ZIP path traversal', async () => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
    zip.file('../outside.xml', '<x/>');
    const malicious = await zip.generateAsync({ type: 'uint8array' });
    await expect(OpcPackage.open(malicious)).rejects.toThrow(/path traversal/);
  });
});

function hash(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}
